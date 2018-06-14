using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using UcbBack.Models;

namespace UcbBack.Logic.ExcelFiles
{
    public class ValidateContractsFile:ValidateExcelFile
    {
        private static Excelcol[] cols = new[] {new Excelcol("id", typeof(double)), new Excelcol("nombre", typeof(string))};
        public ValidateContractsFile(Stream d, string fn)
            : base(cols, d, fn)
        {
        }

        

    }
}