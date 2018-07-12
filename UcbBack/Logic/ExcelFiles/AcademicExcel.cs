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
    public class AcademicExcel : ValidateExcelFile
    {
        private static Excelcol[] cols = new[]
        {
            new Excelcol("Carnet Identidad", typeof(string)), 
            new Excelcol("Nombre Completo", typeof(string)),
            new Excelcol("Tipo empleado", typeof(string)),
            new Excelcol("Periodo académico", typeof(string)),
            new Excelcol("Sigla Asignatura", typeof(string)),
            new Excelcol("Paralelo", typeof(string)),
            new Excelcol("Horas Académicas por semana", typeof(double)),
            new Excelcol("Horas Académicas por mes", typeof(double)),
            new Excelcol("Identificador de Pago", typeof(string)),
            new Excelcol("Categoría de docente", typeof(string)),
            new Excelcol("Costo hora", typeof(double)),
            new Excelcol("Costo mes", typeof(double)),
            new Excelcol("CUNI", typeof(string)),
            new Excelcol("Identificador de dependencia", typeof(string)),
            new Excelcol("PEI-PO", typeof(string)),
            new Excelcol("Codigo Paralelo SAP", typeof(string)),
        };
        private ApplicationDbContext _context;
        private string mes, gestion, segmentoOrigen;

        public AcademicExcel(Stream data, ApplicationDbContext context, string fileName,string mes, string gestion, string segmentoOrigen,int headerin = 3, int sheets = 1, string resultfileName = "Result")
            : base(cols, data, fileName, headerin, sheets, resultfileName)
        {
            this.segmentoOrigen = segmentoOrigen;
            this.gestion = gestion;
            this.mes = mes;
            _context = context;
            isFormatValid();
        }
        public AcademicExcel(string fileName, int headerin = 1)
            : base(cols, fileName, headerin)
        { }

        public override void toDataBase()
        {
            IXLRange UsedRange = wb.Worksheet(1).RangeUsed();

            for (int i = 1 + headerin; i <= UsedRange.LastRow().RowNumber(); i++)
            {
                _context.DistAcademics.Add(ToDistAcademic(i));
            }

            _context.SaveChanges();
        }

        public override bool ValidateFile()
        {
            var connB1 = B1Connection.Instance;
            bool v1 = VerifyPerson(1, 13, 2);
            bool v2 = VerifyColumnValueIn(3, _context.TipoEmpleadoDists.Select(x => x.Name).ToList(), comment: "Este Tipo empleado no es valido.");
            bool v3 = VerifyParalel(cod:16,periodo: 4, sigla:5);
            bool v4 = VerifyColumnValueIn(9, new List<string> { "PA", "PI", "TH" });
            bool v5 = VerifyColumnValueIn(14, _context.Dependencies.Select(m => m.Cod).Distinct().ToList(), comment: "Esta Dependencia no existe en la Base de Datos Nacional.");
            var pei = connB1.getCostCenter(B1Connection.Dimension.PEI, mes: this.mes, gestion: this.gestion).Cast<string>().ToList();
            pei.Add("0");
            bool v6 = VerifyColumnValueIn(15, pei, comment: "Este PEI no existe en SAP.");
            bool v7 = VerifyColumnValueIn(8, new List<string> { "0" }, comment: "Este valor no puede ser 0", notin: true);
            bool v8 = VerifyColumnValueIn(12, new List<string> { "0" }, comment: "Este valor no puede ser 0", notin: true);

            return isValid() && v1 && v2 && v3 && v4 && v5 && v6;
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
            acad.AcademicHoursWeek = strToDecimal(row,7);
            acad.AcademicHoursMonth = strToDecimal(row, 8);
            acad.IdentificadorPago = wb.Worksheet(sheet).Cell(row, 9).Value.ToString();
            acad.CategoriaDocente = wb.Worksheet(sheet).Cell(row, 10).Value.ToString();
            acad.CostoHora = strToDecimal(row, 11);
            acad.CostoMes = strToDecimal(row, 12);
            acad.CUNI = wb.Worksheet(sheet).Cell(row, 13).Value.ToString();
            acad.Dependency = wb.Worksheet(sheet).Cell(row, 14).Value.ToString();
            acad.PEI = wb.Worksheet(sheet).Cell(row, 15).Value.ToString();
            acad.SAPParaleloUnit = wb.Worksheet(sheet).Cell(row, 16).Value.ToString();
            
            acad.Matched = 0;
            acad.Porcentaje = 0.0m;
            acad.mes = this.mes;
            acad.gestion = this.gestion;
            acad.segmentoOrigen = this.segmentoOrigen;
            return acad;
        }
    }
}