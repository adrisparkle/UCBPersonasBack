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
            new Excelcol("Nombre docente", typeof(string)),
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

        public PostgradoExcel(Stream data, ApplicationDbContext context, string fileName, string mes, string gestion, string segmentoOrigen, int headerin = 3, int sheets = 1, string resultfileName = "Result")
            : base(cols, data, fileName, headerin, sheets, resultfileName)
        {
            this.segmentoOrigen = segmentoOrigen;
            this.gestion = gestion;
            this.mes = mes;
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

            bool v1 = VerifyColumnValueIn(3, connB1.getProjects(col: "PrjName").ToList(), comment: "Este Proyecto no existe en SAP.");
            bool v2 = VerifyColumnValueIn(6, _context.Dependencies.Select(m => m.Cod).Distinct().ToList(), comment: "No existe esta dependencia.");
            bool v3 = VerifyColumnValueIn(8, new List<string>{"POST","EC","INV"}, comment: "No existe este tipo de proyecto.");
            bool v4 = VerifyColumnValueIn(9, new List<string> { "PROF", "TG", "REL", "LEC", "REV", "OTR" }, comment: "No existe este tipo de tarea asignada.");
            bool v5 = VerifyColumnValueIn(10, connB1.getCostCenter(B1Connection.Dimension.PEI).Cast<string>().ToList(), comment: "Este PEI no existe en SAP.");
            bool v6 = VerifyColumnValueIn(11, connB1.getCostCenter(B1Connection.Dimension.Periodo).Cast<string>().ToList(), comment: "Este periodo no existe en SAP.");
            bool v7 = VerifyColumnValueIn(12, connB1.getProjects(), comment: "Este proyecto no existe en SAP.");
            bool v8 = VerifyPerson(ci:1,fullname: 2,CUNI:7,date:gestion+"-"+mes+"-01");
            return isValid() && v1 && v2 && v3 && v4 && v5 && v6 && v7 && v8;
        }

        public Dist_Posgrado ToDistDiscounts(int row, int sheet = 1)
        {
            Dist_Posgrado dis = new Dist_Posgrado();
            dis.Id = _context.Database.SqlQuery<int>("SELECT \"rrhh_Dist_Posgrado_sqs\".nextval FROM DUMMY;").ToList()[0];
            dis.Document = wb.Worksheet(sheet).Cell(row, 1).Value.ToString();
            dis.FullName = wb.Worksheet(sheet).Cell(row, 2).Value.ToString();
            dis.ProjectName = wb.Worksheet(sheet).Cell(row, 3).Value.ToString();
            dis.Vesion = (int) strToDecimal(row, 4);
            dis.TotalPagado = strToDecimal(row,5);
            dis.Dependency = wb.Worksheet(sheet).Cell(row, 6).Value.ToString();
            dis.CUNI = wb.Worksheet(sheet).Cell(row, 7).Value.ToString();
            dis.ProjectType = wb.Worksheet(sheet).Cell(row, 8).Value.ToString();
            dis.TipoTarea = wb.Worksheet(sheet).Cell(row, 9).Value.ToString();
            dis.PEI = wb.Worksheet(sheet).Cell(row, 10).Value.ToString();
            dis.Periodo = wb.Worksheet(sheet).Cell(row, 11).Value.ToString();
            dis.ProjectCode = wb.Worksheet(sheet).Cell(row, 12).Value.ToString();

            dis.Porcentaje = 0m;

            dis.mes = this.mes;
            dis.gestion = this.gestion;
            dis.segmentoOrigen = this.segmentoOrigen;
            return dis;
        }
    }
}