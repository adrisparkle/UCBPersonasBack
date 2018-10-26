using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using ClosedXML.Excel;
using Microsoft.Ajax.Utilities;
using UcbBack.Controllers;
using UcbBack.Logic.B1;
using UcbBack.Models;

namespace UcbBack.Logic.ExcelFiles
{
    public class ContractExcel : ValidateExcelFile
    {
        private static Excelcol[] cols = new[]
        {
            new Excelcol("CUNI", typeof(string)),
            new Excelcol("Dependencia", typeof(string)),
            new Excelcol("Cargo", typeof(string)),
            new Excelcol("Descripcion de Cargo", typeof(string)),
            new Excelcol("Dedicacion", typeof(string)),
            new Excelcol("Vinculacion", typeof(string)),
            new Excelcol("Fecha Inicio", typeof(DateTime)),
            new Excelcol("Fecha Fin", typeof(string))
        };

        private static Excelcol[] peopleCols = new[]
        {
            new Excelcol("C.I", typeof(string)), 
            new Excelcol("Exped", typeof(string)),
            new Excelcol("Tipo documento de identificacion", typeof(string)),
            new Excelcol("primer apellido", typeof(string)),
            new Excelcol("segundo apellido", typeof(string)),
            new Excelcol("nombres", typeof(string)),
            new Excelcol("Apellido casada", typeof(string)),
            new Excelcol("Genero", typeof(string)),
            new Excelcol("AFP", typeof(string)),
            new Excelcol("NUA", typeof(string)),
            new Excelcol("fecha nacimiento", typeof(DateTime)),
        };

        private ValidatePerson validate;
        private int Segment;
        private ApplicationDbContext _context;
        public ContractExcel(Stream data, ApplicationDbContext context, string fileName, int Segment ,int headerin = 1, int sheets = 1, string resultfileName = "PayrollResult")
            : base(cols, data, fileName, headerin: headerin, resultfileName: resultfileName, sheets: sheets)
        {
            this.Segment = Segment;
            _context = context;
            validate = new ValidatePerson();
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
                var p = ToContractDetail(i);
                if (p!=null)
                    _context.ContractDetails.Add(p);
            }

            try
            {
                _context.SaveChanges();

            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
        }

        public override bool ValidateFile()
        {
            /*bool v4 = VerifyColumnValueIn(3, _context.Dependencies.Select(m => m.Cod).Distinct().ToList(), comment: "Esta Dependencia no existe en la Base de Datos Nacional.");
            bool v3 = VerifyColumnValueIn(4, _context.Position.Select(m => m.Name).Distinct().ToList(), comment: "Este Cargo no existe en la Base de Datos Nacional.");
            bool v1 = VerifyColumnValueIn(6, new List<string> { "TC", "TH", "MT" });
            bool v2 = VerifyColumnValueIn(7, new List<string> { "Plazo Fijo", "Permanente", "TH" });

            bool v5 = VerifyPerson(ci: 1, CUNI: 2);*/
            return isValid();// && v1 && v2 && v3 && v4 && v5;
        }

        public People ToPeople(int row, int sheet = 1)
        {
            
            People person = new People();
            person.Document = wb.Worksheet(sheet).Cell(row, 1).Value.ToString();

            var p = _context.Person.FirstOrDefault(x => x.Document ==person.Document);
            if (p != null)
                return null;

            person.Id = People.GetNextId(_context);
            person.Ext = wb.Worksheet(sheet).Cell(row, 2).Value.ToString().ToUpper();
            person.TypeDocument = wb.Worksheet(sheet).Cell(row, 3).Value.ToString().ToUpper();
            person.FirstSurName = wb.Worksheet(sheet).Cell(row, 4).Value.ToString().ToUpper();
            person.SecondSurName = wb.Worksheet(sheet).Cell(row, 5).Value.ToString().ToUpper();
            person.Names = wb.Worksheet(sheet).Cell(row, 6).Value.ToString().ToUpper();
            person.MariedSurName = wb.Worksheet(sheet).Cell(row, 7).Value.ToString().ToUpper();
            person.Gender = wb.Worksheet(sheet).Cell(row, 8).Value.ToString().ToUpper();
            person.AFP = wb.Worksheet(sheet).Cell(row, 9).Value.ToString().ToUpper();
            person.NUA = wb.Worksheet(sheet).Cell(row, 10).Value.ToString();
            var date = wb.Worksheet(sheet).Cell(row, 11).Value.ToString();
            person.BirthDate = DateTime.Parse(date);
            person = validate.UcbCode(person);
            return person;
        }

        public ContractDetail ToContractDetail(int row, int sheet = 1)
        {
            try
            {
                ContractDetail person = new ContractDetail();
                person.Id = ContractDetail.GetNextId(_context);
                person.CUNI = wb.Worksheet(sheet).Cell(row, 1).Value.ToString();
                person.PeopleId = _context.Person.FirstOrDefault(p => p.CUNI == person.CUNI).Id;
                var dep = wb.Worksheet(sheet).Cell(row, 2).Value.ToString();
                var depCont = _context.Dependencies.FirstOrDefault(d => d.Cod == dep);
                if (depCont == null)
                {
                    person.DependencyId = 2;
                }
                else
                    person.DependencyId = depCont.Id;

                var pos = wb.Worksheet(sheet).Cell(row, 3).Value.ToString();
                var posCont = _context.Position.FirstOrDefault(p => p.Name == pos);
                if (posCont == null)
                {
                    person.PositionsId = 2;
                }
                else
                    person.PositionsId = posCont.Id;
                person.PositionDescription = wb.Worksheet(sheet).Cell(row, 4).Value.ToString();
                person.Dedication = wb.Worksheet(sheet).Cell(row, 5).Value.ToString();
                person.Linkage = wb.Worksheet(sheet).Cell(row, 6).Value.ToString();
                person.BranchesId = Segment;
                person.StartDate = wb.Worksheet(sheet).Cell(row, 7).GetDateTime();
                person.EndDate = wb.Worksheet(sheet).Cell(row, 8).Value.ToString() == "" ? (DateTime?)null : wb.Worksheet(sheet).Cell(row, 8).GetDateTime();
                 return person;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }

        }
    }
}