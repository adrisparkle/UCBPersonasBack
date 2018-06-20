using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using UcbBack.Models;

namespace UcbBack.Logic.ExcelFiles
{
    public class PayrollExcel:ValidateExcelFile
    {
        private static Excelcol[] cols = new[]
        {
            new Excelcol("Carnet Identidad", typeof(string)), 
            new Excelcol("Nombre Completo", typeof(string)),
            new Excelcol("Haber Basico", typeof(double)),
            new Excelcol("Bono de Antigüedad", typeof(double)),
            new Excelcol("Otros Ingresos", typeof(double)),
            new Excelcol("Ingresos por docencia", typeof(double)),
            new Excelcol("Ingresos por otras actividades academicas", typeof(double)),
            new Excelcol("Total Ganado", typeof(double)),
            new Excelcol("Identificador de AFP", typeof(string)),
            new Excelcol("Aporte Laboral AFP", typeof(double)),
            new Excelcol("RC-IVA", typeof(double)),
            new Excelcol("Descuentos", typeof(double)),
            new Excelcol("Anticipos", typeof(double)),
            new Excelcol("Total Deducciones", typeof(double)),
            new Excelcol("Liquido Pagable", typeof(double)),
            new Excelcol("Segmento", typeof(string)),
            new Excelcol("Codigo nacional RRHH", typeof(string)),
            new Excelcol("Tipo empleado", typeof(string)),
            new Excelcol("PEI", typeof(string)),
            new Excelcol("Horas de trabajo mensual", typeof(double)),
            new Excelcol("Fecha Nacimiento", typeof(string)),
            new Excelcol("Unidad Organigrama", typeof(string)),
            new Excelcol("Aporte Patronal AFP", typeof(double)),
            new Excelcol("Aporte Patronal Seguridad Corto Plazo", typeof(double)),
            new Excelcol("Provision para Aguinaldos", typeof(double)),
            new Excelcol("Provision para Primas", typeof(double)),
            new Excelcol("Provisión para Indeminizacion", typeof(double)),
            new Excelcol("Unidad Organizacional SAP", typeof(string))
        };
        private ApplicationDbContext _context;
        public PayrollExcel(Stream data,ApplicationDbContext context, string fileName, int headerin =1, int sheets = 1, string resultfileName = "PayrollResult")
            : base(cols, data, fileName, headerin: headerin, resultfileName: resultfileName,sheets:sheets)
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
            //bool v1 = VerifyColumnValueIn(1, _context.Person.Select(m => m.Document).Distinct().ToList(),comment:"Este Documento de Identidad no esta registrado en la Base de Datos Nacional.");
           // bool v2 = VerifyColumnValueIn(17, _context.Person.Select(m => m.CUNI).Distinct().ToList(),comment:"Este CUNI no existe en la Base de Datos Nacional.");
            bool v3 = VerifyColumnValueIn(22, _context.Dependencies.Select(m => m.Name).Distinct().ToList(),comment:"Esta Dependencia no existe en la Base de Datos Nacional.");
            bool v4 = VerifyColumnValueIn(28, _context.OrganizationalUnits.Select(m => m.Name).Distinct().ToList(),comment:"Esta Unidad Organigrama no existe en la Base de Datos Nacional.");
            bool v5 = VerifyPerson(1, 17, 2);
            return isValid() && v3 && v4 && v5;
        }
    }
}