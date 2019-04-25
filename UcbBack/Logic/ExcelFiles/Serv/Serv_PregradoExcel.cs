using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using ClosedXML.Excel;
using UcbBack.Models;
using UcbBack.Models.Serv;

namespace UcbBack.Logic.ExcelFiles.Serv
{
    public class Serv_PregradoExcel : ValidateExcelFile
    {
        private static Excelcol[] cols = new[]
        {
            new Excelcol("Codigo Socio de Negocio", typeof(string)), 
            new Excelcol("Nombre Socio de Negocio", typeof(string)),
            new Excelcol("Cod Dependencia", typeof(string)),
            new Excelcol("PEI PO", typeof(string)),
            new Excelcol("Glosa", typeof(string)),
            new Excelcol("Tipo de Servicio", typeof(string)),
            new Excelcol("Importe del Contrato", typeof(double)),
            new Excelcol("Importe Deducción IUE", typeof(double)),
            new Excelcol("Importe Deducción IT", typeof(double)),
            new Excelcol("Monto a Pagar", typeof(double)),
            new Excelcol("Observaciones", typeof(string)),
        };

        private ApplicationDbContext _context;
        private ServProcess process;

        public Serv_PregradoExcel(string fileName, int headerin = 1)
            : base(cols, fileName, headerin)
        { }

        public Serv_PregradoExcel(Stream data, ApplicationDbContext context, string fileName, ServProcess process, int headerin = 1, int sheets = 1, string resultfileName = "Result")
            : base(cols, data, fileName, headerin, sheets, resultfileName, context)
        {
            this.process = process;
            _context = context;
            isFormatValid();
        }

        public override void toDataBase()
        {
            IXLRange UsedRange = wb.Worksheet(1).RangeUsed();

            for (int i = 1 + headerin; i <= UsedRange.LastRow().RowNumber(); i++)
            {
                _context.ServPregrados.Add(ToServVarios(i));
            }

            _context.SaveChanges();
        }

        public Serv_Pregrado ToServVarios(int row, int sheet = 1)
        {
            Serv_Pregrado data = new Serv_Pregrado();
            data.Id = Serv_Varios.GetNextId(_context);

            data.CardCode = wb.Worksheet(sheet).Cell(row, 1).Value.ToString();
            data.CardName = wb.Worksheet(sheet).Cell(row, 2).Value.ToString();
            var cod = wb.Worksheet(sheet).Cell(row, 3).Value.ToString();
            var depId = _context.Dependencies
                .FirstOrDefault(x => x.Cod == cod);
            data.DependencyId = depId.Id;
            data.PEI = wb.Worksheet(sheet).Cell(row, 4).Value.ToString();
            data.Memo = wb.Worksheet(sheet).Cell(row, 5).Value.ToString();
            data.Carrera = wb.Worksheet(sheet).Cell(row, 6).Value.ToString();
            data.DocumentNumber = wb.Worksheet(sheet).Cell(row, 7).Value.ToString();
            data.Student = wb.Worksheet(sheet).Cell(row, 8).Value.ToString();
            data.JobType = wb.Worksheet(sheet).Cell(row, 9).Value.ToString();
            data.Hours = Int32.Parse(wb.Worksheet(sheet).Cell(row, 10).Value.ToString());
            data.CostPerHour = Decimal.Parse(wb.Worksheet(sheet).Cell(row, 11).Value.ToString());
            data.ServiceType = wb.Worksheet(sheet).Cell(row, 6).Value.ToString();
            data.ContractAmount = Decimal.Parse(wb.Worksheet(sheet).Cell(row, 7).Value.ToString());
            data.IUE = Decimal.Parse(wb.Worksheet(sheet).Cell(row, 8).Value.ToString());
            data.IT = Decimal.Parse(wb.Worksheet(sheet).Cell(row, 9).Value.ToString());
            data.TotalAmount = Decimal.Parse(wb.Worksheet(sheet).Cell(row, 10).Value.ToString());
            data.Comments = wb.Worksheet(sheet).Cell(row, 11).Value.ToString();
            data.Serv_ProcessId = process.Id;
            return data;
        }

        public override bool ValidateFile()
        {
            return true;
        }
    }
}