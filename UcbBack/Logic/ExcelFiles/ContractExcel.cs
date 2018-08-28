using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using ClosedXML.Excel;
using UcbBack.Controllers;
using UcbBack.Logic.B1;
using UcbBack.Models;

namespace UcbBack.Logic.ExcelFiles
{
    public class ContractExcel : ValidateExcelFile
    {
        private static Excelcol[] cols = new[]
        {
            new Excelcol("Carnet Identidad", typeof(string)), 
            new Excelcol("CUNI", typeof(string)),
            new Excelcol("Dependencia", typeof(double)),
            new Excelcol("Cargo", typeof(double)),
            new Excelcol("Descripcion de Cargo", typeof(double)),
            new Excelcol("Dedicacion", typeof(double)),
            new Excelcol("Vinculacion", typeof(double)),
            new Excelcol("Fecha Inicio", typeof(DateTime)),
            new Excelcol("Fecha Fin", typeof(DateTime))
        };

        private int Segment;
        private ApplicationDbContext _context;
        public ContractExcel(Stream data, ApplicationDbContext context, string fileName, int Segment ,int headerin = 1, int sheets = 1, string resultfileName = "PayrollResult")
            : base(cols, data, fileName, headerin: headerin, resultfileName: resultfileName, sheets: sheets)
        {
            this.Segment = Segment;
            _context = context;
            isFormatValid();
        }
        public ContractExcel(string fileName, int headerin = 1)
            : base(cols, fileName, headerin)
        { }

        public override void toDataBase()
        {
            IXLRange UsedRange = wb.Worksheet(1).RangeUsed();

            for (int i = 1 + headerin; i <= UsedRange.LastRow().RowNumber(); i++)
            {
                _context.ContractDetails.Add(ToDistPayroll(i));
            }

            _context.SaveChanges();
        }

        public override bool ValidateFile()
        {
            bool v4 = VerifyColumnValueIn(3, _context.Dependencies.Select(m => m.Cod).Distinct().ToList(), comment: "Esta Dependencia no existe en la Base de Datos Nacional.");
            bool v3 = VerifyColumnValueIn(4, _context.Position.Select(m => m.Name).Distinct().ToList(), comment: "Este Cargo no existe en la Base de Datos Nacional.");
            bool v1 = VerifyColumnValueIn(6, new List<string> { "TC", "TH", "MT" });
            bool v2 = VerifyColumnValueIn(7, new List<string> { "Plazo Fijo", "Permanente", "TH" });

            bool v5 = VerifyPerson(ci: 1, CUNI: 2);
            return isValid() && v1 && v2 && v3 && v4 && v5;
        }

        public ContractDetail ToDistPayroll(int row, int sheet = 1)
        {
            ContractDetail person = new ContractDetail();
            person.Id = _context.Database.SqlQuery<int>("SELECT ADMNALRRHH.\"rrhh_ContractDetail_sqs\".nextval FROM DUMMY;").ToList()[0];
            person.CUNI = wb.Worksheet(sheet).Cell(row, 2).Value.ToString();
            person.PeopleId = _context.Person.FirstOrDefault(p => p.CUNI == person.CUNI).Id;
            person.DependencyId = _context.Dependencies.FirstOrDefault(d => d.Cod == wb.Worksheet(sheet).Cell(row, 3).Value.ToString()).Id;
            person.PositionsId = _context.Position.FirstOrDefault(p => p.Name == wb.Worksheet(sheet).Cell(row, 4).Value.ToString()).Id;
            person.PositionDescription = wb.Worksheet(sheet).Cell(row, 5).Value.ToString();
            person.Dedication = wb.Worksheet(sheet).Cell(row, 6).Value.ToString();
            person.Linkage = wb.Worksheet(sheet).Cell(row, 7).Value.ToString();
            person.BranchesId = Segment;
            person.StartDate = wb.Worksheet(sheet).Cell(row, 8).GetDateTime();
            person.EndDate = wb.Worksheet(sheet).Cell(row, 9).GetDateTime();
            return person;
        }
    }
}