using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using UcbBack.Models;

namespace UcbBack.Logic.ExcelFiles
{
    public class Serv_VariosExcel : ValidateExcelFile
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
        private string mes, gestion, segmentoOrigen;

        public Serv_VariosExcel(string fileName, int headerin = 1)
            : base(cols, fileName, headerin)
        { }

        public Serv_VariosExcel(Stream data, ApplicationDbContext context, string fileName, int headerin = 1, int sheets = 1, string resultfileName = "Result") 
            : base(cols, data, fileName, headerin, sheets, resultfileName, context)
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
            throw new NotImplementedException();
        }
    }
}