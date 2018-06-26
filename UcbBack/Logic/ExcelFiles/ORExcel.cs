using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using ClosedXML.Excel;
using UcbBack.Models;

namespace UcbBack.Logic.ExcelFiles
{
    public class ORExcel : ValidateExcelFile
    {
        private static Excelcol[] cols = new[]
        {
            new Excelcol("ID", typeof(string)), 
            new Excelcol("C.I.", typeof(string)),
            new Excelcol("Nombre Completo", typeof(string)),
            new Excelcol("Segmento", typeof(string)),
            new Excelcol("Haber Mensual", typeof(double)),
            new Excelcol("Otros Ingresos", typeof(double)),
            new Excelcol("Total Pagado", typeof(double)),
        };
        private ApplicationDbContext _context;

        public ORExcel(Stream data, ApplicationDbContext context, string fileName, int headerin = 3, int sheets = 1, string resultfileName = "Result")
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
                _context.DistOrs.Add(ToDistDiscounts(i));
            }

            _context.SaveChanges();
        }

        public override bool ValidateFile()
        {
            bool v2 = VerifyPerson(ci:2, fullname:3);
            return isValid()&&v2;
        }

        public Dist_OR ToDistDiscounts(int row, int sheet = 1)
        {
            Dist_OR dis = new Dist_OR();
            dis.Id = _context.Database.SqlQuery<int>("SELECT \"rrhh_Dist_OR_sqs\".nextval FROM DUMMY;").ToList()[0];
            dis.Document = wb.Worksheet(sheet).Cell(row, 2).Value.ToString();
            dis.FullName = wb.Worksheet(sheet).Cell(row, 3).Value.ToString();
            dis.Segment = wb.Worksheet(sheet).Cell(row, 4).Value.ToString();
            dis.HaberMensual = wb.Worksheet(sheet).Cell(row, 5).Value.ToString() == "" ? 0.0m : Decimal.Parse(wb.Worksheet(sheet).Cell(row, 5).Value.ToString());
            dis.OtrosIngresos = wb.Worksheet(sheet).Cell(row, 6).Value.ToString() == "" ? 0.0m : Decimal.Parse(wb.Worksheet(sheet).Cell(row, 6).Value.ToString());
            dis.TotalGanado = wb.Worksheet(sheet).Cell(row, 7).Value.ToString() == "" ? 0.0m : Decimal.Parse(wb.Worksheet(sheet).Cell(row, 7).Value.ToString());

            return dis;
        }
    }
}