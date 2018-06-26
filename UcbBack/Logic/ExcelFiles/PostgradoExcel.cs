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
            new Excelcol("ID", typeof(string)), 
            new Excelcol("C.I.", typeof(string)), 
            new Excelcol("Nombre Completo", typeof(string)),
            new Excelcol("Nombre del proyecto", typeof(string)),
            new Excelcol("Versión", typeof(string)),
            new Excelcol("Total pagado", typeof(double)),
            new Excelcol("Tipo proyecto", typeof(string)),
            new Excelcol("Sub tipo proyecto", typeof(string)),
            new Excelcol("Tipo Tarea", typeof(string)),
            new Excelcol("PEI", typeof(string)),
            new Excelcol("Periodo", typeof(string)),
            new Excelcol("Codigo Proyecto", typeof(string))
        };
        private ApplicationDbContext _context;

        public PostgradoExcel(Stream data, ApplicationDbContext context, string fileName, int headerin = 3, int sheets = 1, string resultfileName = "Result")
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
                _context.DistPosgrados.Add(ToDistDiscounts(i));
            }

            _context.SaveChanges();
        }

        public override bool ValidateFile()
        {
            //todo validate if BussinesParner Exist
            //bool v2 = VerifyPerson(1, 13, 2);
            var connB1 = B1Connection.Instance;

            bool v1 = VerifyColumnValueIn(4, connB1.getProjects(col: "PrjName").ToList(), comment: "Este Proyecto no existe en SAP.");
            bool v2 = VerifyColumnValueIn(12, connB1.getProjects(col: "PrjName").ToList(), comment: "Este Proyecto no existe en SAP.");
            bool v3 = VerifyColumnValueIn(10, connB1.getCostCenter(B1Connection.Dimension.PEI).ToList(), comment: "Este PEI no existe en SAP.");
            bool v4 = VerifyColumnValueIn(11, connB1.getCostCenter(B1Connection.Dimension.Periodo).ToList(), comment: "Este periodo no existe en SAP.");
            bool v5 = VerifyPerson(ci:2,fullname: 3);
            return isValid()&&v1&&v2&&v3&&v4;
        }

        public Dist_Posgrado ToDistDiscounts(int row, int sheet = 1)
        {
            Dist_Posgrado dis = new Dist_Posgrado();
            dis.Id = _context.Database.SqlQuery<int>("SELECT \"rrhh_Dist_Posgrado_sqs\".nextval FROM DUMMY;").ToList()[0];
            dis.Document = wb.Worksheet(sheet).Cell(row, 2).Value.ToString();
            dis.FullName = wb.Worksheet(sheet).Cell(row, 3).Value.ToString();
            dis.ProjectName = wb.Worksheet(sheet).Cell(row, 4).Value.ToString();
            dis.Vesion = wb.Worksheet(sheet).Cell(row, 5).Value.ToString() == "" ? 0 : Int32.Parse(wb.Worksheet(sheet).Cell(row, 5).Value.ToString());
            dis.TotalPagado = wb.Worksheet(sheet).Cell(row, 6).Value.ToString() == "" ? 0.0m : Decimal.Parse(wb.Worksheet(sheet).Cell(row, 6).Value.ToString());
            dis.ProjectType = wb.Worksheet(sheet).Cell(row, 7).Value.ToString();
            dis.ProjectSubType = wb.Worksheet(sheet).Cell(row, 8).Value.ToString();
            dis.TipoTarea = wb.Worksheet(sheet).Cell(row, 9).Value.ToString();
            dis.PEI = wb.Worksheet(sheet).Cell(row, 10).Value.ToString();
            dis.Periodo = wb.Worksheet(sheet).Cell(row, 11).Value.ToString();
            dis.ProjectCode = wb.Worksheet(sheet).Cell(row, 12).Value.ToString();

            return dis;
        }
    }
}