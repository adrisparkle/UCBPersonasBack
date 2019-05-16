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
    public class Serv_CarreraExcel : ValidateExcelFile
    {
        private static Excelcol[] cols = new[]
        {
            new Excelcol("Codigo Socio", typeof(string)), 
            new Excelcol("Nombre Socio", typeof(string)),
            new Excelcol("Cod Dependencia", typeof(string)),
            new Excelcol("PEI PO", typeof(string)),
            new Excelcol("Nombre del Servicio", typeof(string)),
            new Excelcol("Codigo Carrera", typeof(string)),
            new Excelcol("Documento Base", typeof(string)),
            new Excelcol("Postulante", typeof(string)),
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

        public Serv_CarreraExcel(string fileName, int headerin = 1)
            : base(cols, fileName, headerin)
        { }

        public Serv_CarreraExcel(Stream data, ApplicationDbContext context, string fileName, ServProcess process,CustomUser user, int headerin = 1, int sheets = 1, string resultfileName = "Result")
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
                _context.ServCarreras.Add(ToServVarios(i));
            }

            _context.SaveChanges();
        }

        public Serv_Carrera ToServVarios(int row, int sheet = 1)
        {
            Serv_Carrera data = new Serv_Carrera();
            data.Id = Serv_Carrera.GetNextId(_context);

            data.CardCode = wb.Worksheet(sheet).Cell(row, 1).Value.ToString();
            data.CardName = wb.Worksheet(sheet).Cell(row, 2).Value.ToString();
            var cod = wb.Worksheet(sheet).Cell(row, 3).Value.ToString();
            var depId = _context.Dependencies
                .FirstOrDefault(x => x.Cod == cod);
            data.DependencyId = depId.Id;
            data.PEI = wb.Worksheet(sheet).Cell(row, 4).Value.ToString();
            data.ServiceName = wb.Worksheet(sheet).Cell(row, 5).Value.ToString();
            data.Carrera = wb.Worksheet(sheet).Cell(row, 6).Value.ToString();
            data.DocumentNumber = wb.Worksheet(sheet).Cell(row, 7).Value.ToString();
            data.Student = wb.Worksheet(sheet).Cell(row, 8).Value.ToString();
            data.AssignedJob = wb.Worksheet(sheet).Cell(row, 9).Value.ToString();

            data.AssignedAccount = wb.Worksheet(sheet).Cell(row, 10).Value.ToString();
            data.ContractAmount = Decimal.Parse(wb.Worksheet(sheet).Cell(row, 11).Value.ToString());
            data.IUE = Decimal.Parse(wb.Worksheet(sheet).Cell(row, 12).Value.ToString());
            data.IT = Decimal.Parse(wb.Worksheet(sheet).Cell(row, 13).Value.ToString());
            data.TotalAmount = Decimal.Parse(wb.Worksheet(sheet).Cell(row, 14).Value.ToString());
            data.Comments = wb.Worksheet(sheet).Cell(row, 15).Value.ToString();
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
                var pei = connB1.getCostCenter(B1Connection.Dimension.PEI).Cast<string>().ToList();
                bool v3 = VerifyColumnValueIn(4, pei, comment: "Este PEI no existe en SAP.");
                bool v4 = VerifyLength(5, 50);
                var brs = _context.Branch.FirstOrDefault(x => x.Id == process.BranchesId);
                var carrera = connB1.getCostCenter(B1Connection.Dimension.PlanAcademico).Cast<string>().ToList().Where(x => x.Contains(brs.Abr)).ToList();
                bool v5 = VerifyColumnValueIn(6, carrera, comment: "Esta Carrera no existe en SAP.");
                bool v6 = VerifyColumnValueIn(9, new List<string> { "TG", "REL", "LEC", "REV", "PAN", "EXA", "OTR" }, comment: "No existe esta tipo de Tarea Asignada.");
                bool v7 = VerifyColumnValueIn(10, new List<string> { "CC_TEMPORAL" }, comment: "No existe este tipo de Cuenta Asignada.");
                bool v8 = VerifyTotal();

                bool v9 = true;
                foreach (var i in new List<int>() { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14 })
                {
                    v9 = VerifyNotEmpty(i) && v9;
                }

                return v1 && v2 && v3 && v4 && v5 && v6 && v7 && v8 && v9;
            }

            return false;
        }

        private bool VerifyTotal()
        {
            bool res = true;
            int sheet = 1;

            IXLRange UsedRange = wb.Worksheet(sheet).RangeUsed();
            for (int i = headerin + 1; i <= UsedRange.LastRow().RowNumber(); i++)
            {
                decimal contrato = Math.Round(Decimal.Parse(wb.Worksheet(sheet).Cell(i, 11).Value.ToString()), 2);
                decimal IUE = Math.Round(Decimal.Parse(wb.Worksheet(sheet).Cell(i, 12).Value.ToString()), 2);
                decimal IT = Math.Round(Decimal.Parse(wb.Worksheet(sheet).Cell(i, 13).Value.ToString()), 2);
                decimal total = Math.Round(Decimal.Parse(wb.Worksheet(sheet).Cell(i, 14).Value.ToString()), 2);

                if (contrato - IUE - IT != total)
                {
                    res = false;
                    paintXY(11, i, XLColor.Red, "Este valor no cuadra (Contrato - IUE - IT != Monto a Pagar)");
                }
            }

            valid = valid && res;
            if (!res)
                addError("Valor no valido", "Monto a Pagar no cuadra.", false);
            return res;
        }
    }
}