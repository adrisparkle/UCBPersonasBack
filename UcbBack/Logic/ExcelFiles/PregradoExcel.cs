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
    public class PregradoExcel:ValidateExcelFile
    {
        private static Excelcol[] cols = new[]
        {
            new Excelcol("Carnet identidad", typeof(string)), 
            new Excelcol("Nombre Completo", typeof(string)),
            new Excelcol("Carrera", typeof(string)),
            new Excelcol("Total Bruto", typeof(double)),
            new Excelcol("Porcentaje aplicación", typeof(double)),
            new Excelcol("Total Neto", typeof(double))
        };
        private ApplicationDbContext _context;

        public PregradoExcel(Stream data, ApplicationDbContext context, string fileName, int headerin = 3, int sheets = 1, string resultfileName = "Result")
            : base(cols, data, fileName, headerin, sheets, resultfileName)
        {
            _context = context;
            isFormatValid();
        }

        public override void toDataBase()
        {
            IXLRange UsedRange = wb.Worksheet(1).RangeUsed();

            for (int i = 1 + headerin; i <= UsedRange.RowCount() + headerin; i++)
            {
                _context.DistPregrados.Add(ToDistDiscounts(i));
            }

            _context.SaveChanges();
        }

        public override bool ValidateFile()
        {
            var connB1 = B1Connection.Instance;
            var xx = connB1.getCostCenter(B1Connection.Dimension.PlanAcademico).ToList();
            bool v3 = VerifyColumnValueIn(3, connB1.getCostCenter(B1Connection.Dimension.PlanAcademico).ToList(), comment: "Este Plan de Estudio no existe en SAP.");
            bool v5 = VerifyPerson(ci:1, fullname: 2);

            return isValid()&&v3&&v5;
        }

        public Dist_Pregrado ToDistDiscounts(int row, int sheet = 1)
        {
            Dist_Pregrado dis = new Dist_Pregrado();
            dis.Id = _context.Database.SqlQuery<int>("SELECT \"rrhh_Dist_Pregrado_sqs\".nextval FROM DUMMY;").ToList()[0];
            dis.Document = wb.Worksheet(sheet).Cell(row, 1).Value.ToString();
            dis.FullName = wb.Worksheet(sheet).Cell(row, 2).Value.ToString();
            dis.Carrera = wb.Worksheet(sheet).Cell(row, 3).Value.ToString();
            dis.TotalBruto = wb.Worksheet(sheet).Cell(row, 4).Value.ToString() == "" ? 0.0m : Decimal.Parse(wb.Worksheet(sheet).Cell(row, 3).Value.ToString());
            dis.Porcentaje = wb.Worksheet(sheet).Cell(row, 5).Value.ToString() == "" ? 0.0m : Decimal.Parse(wb.Worksheet(sheet).Cell(row, 4).Value.ToString());
            dis.TotalNeto = wb.Worksheet(sheet).Cell(row, 6).Value.ToString() == "" ? 0.0m : Decimal.Parse(wb.Worksheet(sheet).Cell(row, 5).Value.ToString());

            return dis;
        }
    }
}