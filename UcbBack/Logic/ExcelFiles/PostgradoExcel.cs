using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using ClosedXML.Excel;
using UcbBack.Logic.B1;
using UcbBack.Models;

namespace UcbBack.Logic.ExcelFiles
{
    public class PostgradoExcel : ValidateExcelFile
    {
        private static Excelcol[] cols = new[]
        {
            new Excelcol("Carnet Identidad", typeof(string)), 
            new Excelcol("Primer Apellido", typeof(string)),
            new Excelcol("Segundo Apellido", typeof(string)),
            new Excelcol("Nombres", typeof(string)),
            new Excelcol("Apellido Casada", typeof(string)),
            new Excelcol("Nombre del Proyecto", typeof(string)),
            new Excelcol("Versión", typeof(string)),
            new Excelcol("Total Neto Ganado", typeof(double)),
            new Excelcol("Identificador de Dependencia", typeof(string)),
            new Excelcol("CUNI", typeof(string)),
            new Excelcol("Tipo Proyecto", typeof(string)),
            new Excelcol("Tipo de tarea asignada", typeof(string)),
            new Excelcol("PEI", typeof(string)),
            new Excelcol("Periodo académico", typeof(string)),
            new Excelcol("Código Proyecto SAP", typeof(string))
        };
        private ApplicationDbContext _context;
        private string mes, gestion, segmentoOrigen;
        private Dist_File file;
        public PostgradoExcel(Stream data, ApplicationDbContext context, string fileName, string mes, string gestion, string segmentoOrigen, Dist_File file, int headerin = 3, int sheets = 1, string resultfileName = "Result")
            : base(cols, data, fileName, headerin, sheets, resultfileName)
        {
            this.segmentoOrigen = segmentoOrigen;
            this.gestion = gestion;
            this.mes = mes;
            this.file = file;
            _context = context;
            isFormatValid();
        }
        public PostgradoExcel(string fileName, int headerin = 1)
            : base(cols, fileName, headerin)
        { }

        public override void toDataBase()
        {
            IXLRange UsedRange = wb.Worksheet(1).RangeUsed();

            for (int i = 1 + headerin; i <= UsedRange.LastRow().RowNumber(); i++)
            {
                _context.DistPosgrados.Add(ToDistDiscounts(i));
            }

            _context.SaveChanges();
        }

        public override bool ValidateFile()
        {
            var connB1 = B1Connection.Instance;

            //bool v1 = VerifyColumnValueIn(6, connB1.getProjects(col: "PrjName").ToList(), comment: "Este Proyecto no existe en SAP.");
            bool v2 = VerifyColumnValueIn(9, _context.Dependencies.Select(m => m.Cod).Distinct().ToList(), comment: "No existe esta dependencia.");
            bool v3 = VerifyColumnValueIn(11, new List<string>{"POST","EC","INV"}, comment: "No existe este tipo de proyecto.");
            bool v4 = VerifyColumnValueIn(12, new List<string> { "PROF", "TG", "REL", "LEC", "REV", "OTR", "PAN" }, comment: "No existe este tipo de tarea asignada.");
            //bool v5 = VerifyColumnValueIn(13, connB1.getCostCenter(B1Connection.Dimension.PEI, mes: this.mes, gestion: this.gestion).Cast<string>().ToList(), comment: "Este PEI no existe en SAP.");
            bool v6 = VerifyColumnValueIn(14, connB1.getCostCenter(B1Connection.Dimension.Periodo, mes: this.mes, gestion: this.gestion).Cast<string>().ToList(), comment: "Este periodo no existe en SAP.");
            //bool v7 = VerifyColumnValueIn(15, connB1.getProjects(), comment: "Este proyecto no existe en SAP.");
           // bool v7 = verifyproject();
            bool v8 = VerifyPerson(ci: 1, fullname: 2, CUNI: 10, date: gestion + "-" + mes + "-01", personActive: false);
            return isValid()  && v2 && v3 && v4&& v8 /*&&v5 && v7*/;
        }

        private bool verifyproject(int sheet = 1)
        {
            string commnet = "Este proyecto no existe en SAP.";
            var connB1 = B1Connection.Instance;
            List<string> list = connB1.getProjects();
            int index = 15;
            int tipoproy = 11;
            bool res = true;
            IXLRange UsedRange = wb.Worksheet(sheet).RangeUsed();
            var l = UsedRange.LastRow().RowNumber();
            for (int i = headerin + 1; i <= UsedRange.LastRow().RowNumber(); i++)
            {
                if (!list.Exists(x => string.Equals(x, wb.Worksheet(sheet).Cell(i, index).Value.ToString(), StringComparison.OrdinalIgnoreCase)))
                {
                    var a1 = wb.Worksheet(sheet).Cell(i, tipoproy).Value.ToString();
                    var a2 = wb.Worksheet(sheet).Cell(i, index).Value.ToString();
                    if (!(wb.Worksheet(sheet).Cell(i, tipoproy).Value.ToString() == "EC" &&
                        wb.Worksheet(sheet).Cell(i, index).Value.ToString() == ""))
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

        public Dist_Posgrado ToDistDiscounts(int row, int sheet = 1)
        {
            Dist_Posgrado dis = new Dist_Posgrado();
            dis.Id = _context.Database.SqlQuery<int>("SELECT ADMNALRRHH.\"rrhh_Dist_Posgrado_sqs\".nextval FROM DUMMY;").ToList()[0];
            dis.Document = wb.Worksheet(sheet).Cell(row, 1).Value.ToString();
            dis.FirstName = wb.Worksheet(sheet).Cell(row, 2).Value.ToString();
            dis.FirstSurName = wb.Worksheet(sheet).Cell(row, 3).Value.ToString();
            dis.SecondSurName = wb.Worksheet(sheet).Cell(row, 4).Value.ToString();
            dis.MariedSurName = wb.Worksheet(sheet).Cell(row, 5).Value.ToString();
            dis.ProjectName = wb.Worksheet(sheet).Cell(row, 6).Value.ToString();
            dis.Vesion = (int) strToDecimal(row, 7);
            dis.TotalPagado = strToDecimal(row,8);
            dis.Dependency = wb.Worksheet(sheet).Cell(row, 9).Value.ToString();
            dis.CUNI = wb.Worksheet(sheet).Cell(row, 10).Value.ToString();
            dis.ProjectType = wb.Worksheet(sheet).Cell(row, 11).Value.ToString();
            dis.TipoTarea = wb.Worksheet(sheet).Cell(row, 12).Value.ToString();
            dis.PEI = wb.Worksheet(sheet).Cell(row, 13).Value.ToString();
            dis.Periodo = wb.Worksheet(sheet).Cell(row, 14).Value.ToString();
            dis.ProjectCode = wb.Worksheet(sheet).Cell(row, 15).Value.ToString();

            dis.Porcentaje = 0m;

            dis.mes = this.mes;
            dis.gestion = this.gestion;
            dis.segmentoOrigen = this.segmentoOrigen;

            dis.DistFileId = file.Id;
            return dis;
        }
    }
}