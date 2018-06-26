using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using ClosedXML.Excel;
using UcbBack.Models;

namespace UcbBack.Logic.ExcelFiles
{
    public class AcademicExcel : ValidateExcelFile
    {
        private static Excelcol[] cols = new[]
        {
            new Excelcol("Carnet Identidad", typeof(string)), 
            new Excelcol("Nombre Completo", typeof(string)),
            new Excelcol("Tipo empleado", typeof(string)),
            new Excelcol("periodo academico", typeof(string)),
            new Excelcol("Sigla Asignatura", typeof(string)),
            new Excelcol("Paralelo", typeof(string)),
            new Excelcol("Horas Academicas por semana", typeof(double)),
            new Excelcol("Horas Academicas por mes", typeof(double)),
            new Excelcol("Identificador de Pago", typeof(string)),
            new Excelcol("Categoria docente", typeof(string)),
            new Excelcol("Costo hora", typeof(double)),
            new Excelcol("Costo mes", typeof(double)),
            new Excelcol("Codigo nacional RRHH", typeof(string)),
            new Excelcol("Codigo Paralelo SAP", typeof(string)),
            new Excelcol("Porcentaje", typeof(double)),
        };
        private ApplicationDbContext _context;

        public AcademicExcel(Stream data, ApplicationDbContext context, string fileName, int headerin = 3, int sheets = 1, string resultfileName = "Result")
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
                _context.DistAcademics.Add(ToDistAcademic(i));
            }

            _context.SaveChanges();
        }

        public override bool ValidateFile()
        {
            bool v2 = VerifyPerson(1, 13, 2);
            return isValid() && v2;
        }

        public Dist_Academic ToDistAcademic(int row,int sheet = 1)
        {
            Dist_Academic acad = new Dist_Academic();
            acad.Id = _context.Database.SqlQuery<int>("SELECT \"rrhh_Dist_Academic_sqs\".nextval FROM DUMMY;").ToList()[0];
            acad.Document = wb.Worksheet(sheet).Cell(row, 1).Value.ToString();
            acad.FullName = wb.Worksheet(sheet).Cell(row, 2).Value.ToString();
            acad.EmployeeType = wb.Worksheet(sheet).Cell(row, 3).Value.ToString();
            acad.Periodo = wb.Worksheet(sheet).Cell(row, 4).Value.ToString();
            acad.Sigla = wb.Worksheet(sheet).Cell(row, 5).Value.ToString();
            acad.Paralelo = wb.Worksheet(sheet).Cell(row, 6).Value.ToString();
            acad.AcademicHoursWeek = wb.Worksheet(sheet).Cell(row, 7).Value.ToString() == "" ? 0.0m : Decimal.Parse(wb.Worksheet(sheet).Cell(row, 7).Value.ToString());
            acad.AcademicHoursMonth = wb.Worksheet(sheet).Cell(row, 8).Value.ToString() == "" ? 0.0m : Decimal.Parse(wb.Worksheet(sheet).Cell(row, 8).Value.ToString());
            acad.IdentificadorPago = wb.Worksheet(sheet).Cell(row, 9).Value.ToString();
            acad.CategoriaDocuente = wb.Worksheet(sheet).Cell(row, 10).Value.ToString();
            acad.CostoHora = wb.Worksheet(sheet).Cell(row, 11).Value.ToString() == "" ? 0.0m : Decimal.Parse(wb.Worksheet(sheet).Cell(row, 11).Value.ToString());
            acad.CostoMes = wb.Worksheet(sheet).Cell(row, 12).Value.ToString() == "" ? 0.0m : Decimal.Parse(wb.Worksheet(sheet).Cell(row, 12).Value.ToString());
            acad.CUNI = wb.Worksheet(sheet).Cell(row, 13).Value.ToString();
            acad.SAPParaleloUnit = wb.Worksheet(sheet).Cell(row, 14).Value.ToString();
            acad.Porcentaje = wb.Worksheet(sheet).Cell(row, 15).Value.ToString()==""? 0.0m : Decimal.Parse(wb.Worksheet(sheet).Cell(row, 15).Value.ToString());
            acad.Matched = 0;
            acad.ProcedureTypeEmployee = "";

            return acad;
        }
    }
}