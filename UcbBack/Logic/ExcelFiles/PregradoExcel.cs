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
    public class PregradoExcel:ValidateExcelFile
    {
        private static Excelcol[] cols = new[]
        {
            new Excelcol("Carnet Identidad", typeof(string)), 
            new Excelcol("Primer Apellido", typeof(string)),
            new Excelcol("Segundo Apellido", typeof(string)),
            new Excelcol("Nombres", typeof(string)),
            new Excelcol("Apellido Casada", typeof(string)),
            new Excelcol("Total Neto Ganado", typeof(double)),
            new Excelcol("Código de Carrera", typeof(string)),
            new Excelcol("CUNI", typeof(string)),
            new Excelcol("Identificador de dependencia", typeof(string)),
        };
        private ApplicationDbContext _context;
        private string mes, gestion, segmentoOrigen;
        private Dist_File file;
        public PregradoExcel(Stream data, ApplicationDbContext context, string fileName, string mes, string gestion, string segmentoOrigen,Dist_File file,int headerin = 3, int sheets = 1, string resultfileName = "Result")
            : base(cols, data, fileName, headerin, sheets, resultfileName)
        {
            this.segmentoOrigen = segmentoOrigen;
            this.gestion = gestion;
            this.mes = mes;
            this.file = file;
            _context = context;
            isFormatValid();
        }

        public PregradoExcel(string fileName, int headerin = 1)
            : base(cols, fileName, headerin)
        { }

        public override void toDataBase()
        {
            IXLRange UsedRange = wb.Worksheet(1).RangeUsed();

            for (int i = 1 + headerin; i <= UsedRange.LastRow().RowNumber(); i++)
            {
                _context.DistPregrados.Add(ToDistDiscounts(i));
            }

            _context.SaveChanges();
        }

        public override bool ValidateFile()
        {
            var connB1 = B1Connection.Instance();
            bool v1 = VerifyPerson(ci: 1, fullname: 2, CUNI: 8, date: this.gestion + "-" + this.mes + "-01", personActive: false);
            bool v2 = VerifyColumnValueIn(7, connB1.getCostCenter(B1Connection.Dimension.PlanAcademico, mes: this.mes, gestion: this.gestion).Cast<string>().ToList(), comment: "Este Plan de Estudio no existe en SAP.");
            bool v3 = VerifyColumnValueIn(9, _context.Dependencies.Select(m => m.Cod).Distinct().ToList());
            return isValid() && v1 && v2 && v3;
        }

        public Dist_Pregrado ToDistDiscounts(int row, int sheet = 1)
        {
            Dist_Pregrado dis = new Dist_Pregrado();
            dis.Id = _context.Database.SqlQuery<int>("SELECT ADMNALRRHH.\"rrhh_Dist_Pregrado_sqs\".nextval FROM DUMMY;").ToList()[0];
            dis.Document = wb.Worksheet(sheet).Cell(row, 1).Value.ToString();
            dis.FirstName = wb.Worksheet(sheet).Cell(row, 2).Value.ToString();
            dis.FirstSurName = wb.Worksheet(sheet).Cell(row, 3).Value.ToString();
            dis.SecondSurName = wb.Worksheet(sheet).Cell(row, 4).Value.ToString();
            dis.MariedSurName = wb.Worksheet(sheet).Cell(row, 5).Value.ToString();
            dis.TotalNeto = strToDecimal(row, 6);
            dis.Carrera = wb.Worksheet(sheet).Cell(row, 7).Value.ToString();
            dis.CUNI = wb.Worksheet(sheet).Cell(row, 8).Value.ToString();
            dis.Dependency = wb.Worksheet(sheet).Cell(row, 9).Value.ToString();

            dis.Porcentaje = 0m;
            dis.mes = this.mes;
            dis.gestion = this.gestion;
            dis.segmentoOrigen = this.segmentoOrigen;

            dis.DistFileId = file.Id;
            return dis;
        }
    }
}