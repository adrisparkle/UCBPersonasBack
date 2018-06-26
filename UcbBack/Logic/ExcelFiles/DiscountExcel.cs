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
    public class DiscountExcel : ValidateExcelFile
    {
        private static Excelcol[] cols = new[]
        {
            new Excelcol("Socio de Negocio", typeof(string)), 
            new Excelcol("Aclaración del Socio de negocio", typeof(string)),
            new Excelcol("Tipo", typeof(string)),
            new Excelcol("Importe", typeof(string)),
        };
        private ApplicationDbContext _context;

        public DiscountExcel(Stream data, ApplicationDbContext context, string fileName, int headerin = 3, int sheets = 1, string resultfileName = "Result")
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
                _context.DistDiscountses.Add(ToDistDiscounts(i));
            }

            _context.SaveChanges();
        }

        public override bool ValidateFile()
        {
            var connB1 = B1Connection.Instance;
            bool v3 = VerifyColumnValueIn(1, connB1.getBusinessPartners().ToList(), comment: "Este Codigo de Socio de Negocio no existe en SAP.");
            bool v4 = VerifyColumnValueIn(2, connB1.getBusinessPartners(col: "CardName").ToList(), comment: "Este nombre de Socio de Negocio no existe en SAP.");

            return isValid()&&v3&&v4;
        }

        public Dist_Discounts ToDistDiscounts(int row, int sheet = 1)
        {
            Dist_Discounts dis = new Dist_Discounts();
            dis.Id = _context.Database.SqlQuery<int>("SELECT \"rrhh_Dist_Discounts_sqs\".nextval FROM DUMMY;").ToList()[0];
            dis.BussinesPartner = wb.Worksheet(sheet).Cell(row, 1).Value.ToString();
            dis.Name = wb.Worksheet(sheet).Cell(row, 2).Value.ToString();
            dis.Type = wb.Worksheet(sheet).Cell(row, 3).Value.ToString();
            dis.Total = wb.Worksheet(sheet).Cell(row, 4).Value.ToString() == "" ? 0.0m : Decimal.Parse(wb.Worksheet(sheet).Cell(row, 4).Value.ToString());

            return dis;
        }
    }
}