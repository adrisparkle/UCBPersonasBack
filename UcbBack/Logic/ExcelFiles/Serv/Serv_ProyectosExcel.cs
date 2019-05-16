using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using ClosedXML.Excel;
using UcbBack.Logic.B1;
using UcbBack.Models;
using UcbBack.Models.Auth;
using UcbBack.Models.Serv;

namespace UcbBack.Logic.ExcelFiles.Serv
{
    public class Serv_ProyectosExcel : ValidateExcelFile
    {
        private static Excelcol[] cols = new[]
        {
            new Excelcol("Codigo Socio", typeof(string)), 
            new Excelcol("Nombre Socio", typeof(string)),
            new Excelcol("Cod Dependencia", typeof(string)),
            new Excelcol("PEI PO", typeof(string)),
            new Excelcol("Nombre del Servicio", typeof(string)),
            new Excelcol("Codigo Proyecto SAP", typeof(string)),
            new Excelcol("Nombre del Proyecto", typeof(string)),
            new Excelcol("Version", typeof(string)),
            new Excelcol("Periodo Academico", typeof(string)),
            new Excelcol("Tipo Tarea Asignada", typeof(string)),
            new Excelcol("Cuenta Asignada", typeof(string)),
            new Excelcol("Monto Contrato", typeof(double)),
            new Excelcol("Monto IUE", typeof(double)),
            new Excelcol("Monto IT", typeof(double)),
            new Excelcol("Monto a Pagar", typeof(double)),
            new Excelcol("Observaciones", typeof(string)),
        };

        private ApplicationDbContext _context;
        private ServProcess process;
        private CustomUser user;

        public Serv_ProyectosExcel(string fileName, int headerin = 1)
            : base(cols, fileName, headerin)
        { }

        public Serv_ProyectosExcel(Stream data, ApplicationDbContext context, string fileName, ServProcess process, CustomUser user,int headerin = 1, int sheets = 1, string resultfileName = "Result")
            : base(cols, data, fileName, headerin, sheets, resultfileName, context)
        {
            this.process = process;
            this.user = user;
            _context = context;
            isFormatValid();
        }

        public override void toDataBase()
        {
            IXLRange UsedRange = wb.Worksheet(1).RangeUsed();

            for (int i = 1 + headerin; i <= UsedRange.LastRow().RowNumber(); i++)
            {
                _context.ServProyectoses.Add(ToServVarios(i));
            }

            _context.SaveChanges();
        }

        public Serv_Proyectos ToServVarios(int row, int sheet = 1)
        {
            Serv_Proyectos data = new Serv_Proyectos();
            data.Id = Serv_Proyectos.GetNextId(_context);

            data.CardCode = wb.Worksheet(sheet).Cell(row, 1).Value.ToString();
            data.CardName = wb.Worksheet(sheet).Cell(row, 2).Value.ToString();
            var cod = wb.Worksheet(sheet).Cell(row, 3).Value.ToString();
            var depId = _context.Dependencies
                .FirstOrDefault(x => x.Cod == cod);
            data.DependencyId = depId.Id;
            data.PEI = wb.Worksheet(sheet).Cell(row, 4).Value.ToString();
            data.ServiceName = wb.Worksheet(sheet).Cell(row, 5).Value.ToString();
            data.ProjectSAPCode = wb.Worksheet(sheet).Cell(row, 6).Value.ToString();
            data.ProjectSAPName = wb.Worksheet(sheet).Cell(row, 7).Value.ToString();
            data.Version = wb.Worksheet(sheet).Cell(row, 8).Value.ToString();
            data.Periodo = wb.Worksheet(sheet).Cell(row, 9).Value.ToString();
            data.AssignedJob = wb.Worksheet(sheet).Cell(row, 10).Value.ToString();

            data.AssignedAccount = wb.Worksheet(sheet).Cell(row, 11).Value.ToString();
            data.ContractAmount = Decimal.Parse(wb.Worksheet(sheet).Cell(row, 12).Value.ToString());
            data.IUE = Decimal.Parse(wb.Worksheet(sheet).Cell(row, 13).Value.ToString());
            data.IT = Decimal.Parse(wb.Worksheet(sheet).Cell(row, 14).Value.ToString());
            data.TotalAmount = Decimal.Parse(wb.Worksheet(sheet).Cell(row, 15).Value.ToString());
            data.Comments = wb.Worksheet(sheet).Cell(row, 16).Value.ToString();
            data.Serv_ProcessId = process.Id;
            return data;
        }

        public override bool ValidateFile()
        {
            if (isValid())
            {
                var connB1 = B1Connection.Instance();

                if (!connB1.connectedtoHana)
                {
                    addError("Error en SAP", "No se puedo conectar con SAP B1, es posible que algunas validaciones cruzadas con SAP no sean ejecutadas");
                }

                bool v1 = VerifyBP(1, 2,process.BranchesId,user);
                bool v2 = VerifyColumnValueIn(3, _context.Dependencies.Where(x => x.BranchesId == this.process.BranchesId).Select(x => x.Cod).ToList(), comment: "Esta Dependencia no es Válida");
                var pei = connB1.getCostCenter(B1Connection.Dimension.PEI).Cast<String>().ToList();
                bool v3 = VerifyColumnValueIn(4, pei, comment: "Este PEI no existe en SAP.");
                bool v4 = VerifyLength(5, 50);
                bool v5 = verifyproject();

                var periodo = connB1.getCostCenter(B1Connection.Dimension.Periodo).Cast<string>().ToList();
                bool v6 = VerifyColumnValueIn(9, periodo, comment: "Este Periodo no existe en SAP.");

                bool v7 = VerifyColumnValueIn(10, new List<string> { "PROF", "TG", "REL", "LEC", "REV", "PAN", "OTR" }, comment: "No existe este tipo de Tarea Asignada.");
                bool v8 = VerifyColumnValueIn(11, new List<string> { "CC_POST", "CC_EC", "CC_FC", "CC_INV", "CC_SA" }, comment: "No existe este tipo de Cuenta Asignada.");

                bool v9 = true;
                foreach (var i in new List<int>(){1,2,3,4,5  ,7  ,9,10,11,12,13,14,15})
                {
                    v9 = VerifyNotEmpty(i) && v9;
                }

                return v1 && v2 && v3 && v4 && v5 && v6 && v7 && v8 && v9;
            }

            return false;

        }

        private bool verifyproject(int sheet = 1)
        {
            string commnet = "Este proyecto no existe en SAP.";
            var connB1 = B1Connection.Instance();
            var br = _context.Branch.FirstOrDefault(x => x.Id == process.BranchesId);
            var list = connB1.getProjects("*").Where(x => x.U_Sucursal == br.Abr).Select(x => x.PrjCode).ToList();
            int index = 6;
            int tipoproy = 11;
            bool res = true;
            IXLRange UsedRange = wb.Worksheet(sheet).RangeUsed();
            var l = UsedRange.LastRow().RowNumber();
            for (int i = headerin + 1; i <= UsedRange.LastRow().RowNumber(); i++)
            {
                if (!list.Exists(x => string.Equals(x.ToString(), wb.Worksheet(sheet).Cell(i, index).Value.ToString(), StringComparison.OrdinalIgnoreCase)))
                {
                    var a1 = wb.Worksheet(sheet).Cell(i, tipoproy).Value.ToString();
                    var a2 = wb.Worksheet(sheet).Cell(i, index).Value.ToString();
                    if (!(
                        (
                            wb.Worksheet(sheet).Cell(i, tipoproy).Value.ToString() == "CC_EC"
                            || wb.Worksheet(sheet).Cell(i, tipoproy).Value.ToString() == "CC_FC"
                            || wb.Worksheet(sheet).Cell(i, tipoproy).Value.ToString() == "CC_SA"
                        )
                        &&
                        wb.Worksheet(sheet).Cell(i, index).Value.ToString() == ""
                    ))
                    {
                        res = false;
                        paintXY(index, i, XLColor.Red, commnet);
                    }
                }
            }
            valid = valid && res;
            if (!res)
                addError("Valor no valido", "Valor o valores no validos en la columna: " + index, false);
            return res;
        }
    }
}