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

        private static Excelcol[] AltaCols = new[]
        {
            new Excelcol("Documento", typeof(string)),
            new Excelcol("Expedido", typeof(string)),
            new Excelcol("Tipo documento de identificacion", typeof(string)),
            new Excelcol("Primer Apellido", typeof(string)),
            new Excelcol("Segundo Apellido", typeof(string)),
            new Excelcol("Nombres", typeof(string)),
            new Excelcol("Apellido casada", typeof(string)),
            new Excelcol("Genero", typeof(string)),
            new Excelcol("AFP", typeof(string)),
            new Excelcol("NUA", typeof(string)),
            new Excelcol("Fecha Nacimiento", typeof(DateTime)),
            new Excelcol("Dependencia", typeof(string)),
            // new Excelcol("Cargo", typeof(string)), default DTH
            // new Excelcol("Descripcion de Cargo", typeof(string)), default DTH
            // new Excelcol("Dedicacion", typeof(string)), default DTH
            // new Excelcol("Vinculacion", typeof(string)), default DTH
            new Excelcol("Fecha Inicio", typeof(DateTime)),
            new Excelcol("Fecha Fin", typeof(DateTime))
        };

        private ValidatePerson validate;
        private int Segment;
        private ApplicationDbContext _context;

        public ContractExcel(Stream data, ApplicationDbContext context, string fileName, int Segment, int headerin = 1,
            int sheets = 1, string resultfileName = "PayrollResult")
            : base(AltaCols, data, fileName, headerin: headerin, resultfileName: resultfileName, sheets: sheets)
        {
            this.Segment = Segment;
            _context = context;
            validate = new ValidatePerson();
            isFormatValid();
        }

        public ContractExcel(string fileName, int headerin = 1)
            : base(AltaCols, fileName, headerin)
        {
        }

        public override void toDataBase()
        {
            IXLRange UsedRange = wb.Worksheet(1).RangeUsed();

            for (int i = 1 + headerin; i <= UsedRange.LastRow().RowNumber(); i++)
            {
                var TempAlta = ToTempAlta(i);

                _context.TempAltas.Add(TempAlta);
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
            bool v1 = VerifyColumnValueIn(2, new List<string> {"LP", "CB", "SC", "TJ", "OR", "CH", "BN", "PA", "PT"});
            bool v2 = VerifyColumnValueIn(3, new List<string> {"CI", "CE", "PA"});
            bool v3 = VerifyColumnValueIn(8, new List<string> {"M", "F"});
            bool v4 = VerifyColumnValueIn(9, new List<string> {"FUT", "PREV"});
            bool v5 = VerifyColumnValueIn(12,
                _context.Dependencies.Where(x => x.BranchesId == this.Segment).Select(m => m.Cod).Distinct().ToList(),
                comment:
                "Esta Dependencia no existe en la Base de Datos Nacional, o no tiene acceso para asociar un empleado a la misma.");
            // todo validations in dates

            return isValid() && v1 && v2 && v3 && v4 && v5;
        }

        public TempAlta ToTempAlta(int row, int sheet = 1)
        {

            TempAlta alta = new TempAlta();
            alta.Document = wb.Worksheet(sheet).Cell(row, 1).Value.ToString();

            var p = _context.Person.FirstOrDefault(x => x.Document == alta.Document);
            if (p == null)
                alta.State = "NEW";
            else
            {
                alta.State = "EXISTING";
                alta.CUNI = p.CUNI;
            }

            alta.Id = TempAlta.GetNextId(_context);
            alta.Ext = wb.Worksheet(sheet).Cell(row, 2).Value.ToString().ToUpper();
            alta.TypeDocument = wb.Worksheet(sheet).Cell(row, 3).Value.ToString().ToUpper();
            alta.FirstSurName = wb.Worksheet(sheet).Cell(row, 4).Value.ToString().ToUpper();
            alta.SecondSurName = wb.Worksheet(sheet).Cell(row, 5).Value.ToString().ToUpper();
            alta.Names = wb.Worksheet(sheet).Cell(row, 6).Value.ToString().ToUpper();
            alta.MariedSurName = wb.Worksheet(sheet).Cell(row, 7).Value.ToString().ToUpper();
            alta.Gender = wb.Worksheet(sheet).Cell(row, 8).Value.ToString().ToUpper();
            alta.AFP = wb.Worksheet(sheet).Cell(row, 9).Value.ToString().ToUpper();
            alta.NUA = wb.Worksheet(sheet).Cell(row, 10).Value.ToString();
            alta.BirthDate = wb.Worksheet(sheet).Cell(row, 11).GetDateTime();
            alta.Dependencia = wb.Worksheet(sheet).Cell(row, 12).Value.ToString();
            alta.StartDate = wb.Worksheet(sheet).Cell(row, 13).GetDateTime();
            alta.EndDate = wb.Worksheet(sheet).Cell(row, 14).GetDateTime();
            alta.BranchesId = Segment;

            return alta;
        }
    };
}