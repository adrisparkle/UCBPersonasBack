using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
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

        public AcademicExcel(Stream data, ApplicationDbContext context, string fileName, int headerin = 1, int sheets = 1, string resultfileName = "Result")
            : base(cols, data, fileName, headerin, sheets, resultfileName)
        {
            _context = context;
            isFormatValid();
        }

        public override void toDataBase()
        {
            throw new NotImplementedException();
        }

        public override bool ValidateFile()
        {
            bool v2 = VerifyPerson(1, 13, 2);
            return isValid() && v2;
        }
    }
}