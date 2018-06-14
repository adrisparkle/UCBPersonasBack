using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using ExcelDataReader;

namespace UcbBack.Logic
{
    public struct Excelcol
    {
        public string headers;
        public Type typeofcol;
        public Excelcol(string h, Type t)
        {
            headers = h;
            typeofcol = t;
        }
    }
    public class ValidateExcelFile
    {
        public Excelcol[] columns { get; set; }
        public DataTable data { get; set; }
        public string fileName { get; set; }
        public ValidateExcelFile(Excelcol[] cols, Stream d,string fn)
        {
            columns = cols;
            data = setExcelFile(d);
            fileName = fn;
        }

        public DataTable setExcelFile(Stream stream)
        {
            IExcelDataReader reader = null;
            if (fileName.EndsWith(".xls"))
            {
                //reads the excel file with .xls extension
                reader = ExcelReaderFactory.CreateBinaryReader(stream);
            }
            else if (fileName.EndsWith(".xlsx"))
            {
                //reads excel file with .xlsx extension
                reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            }
            //Adding reader data to DataSet()
            DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
            {
                ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                {
                    UseHeaderRow = true
                }
            });
            reader.Close();

            foreach (DataTable sheet in result.Tables)
            {
                return sheet;
            }
            return null;
        }

        public bool isFormatValid()
        {
            if (data == null || data.Columns.Count != columns.Length)
            {
                return false;
            }
            for (int i = 0; i < columns.Length; i++)
            {
                if (!String.Equals(data.Columns[i].ColumnName.Trim(), columns[i].headers.Trim(), StringComparison.OrdinalIgnoreCase)
                                || data.Columns[i].DataType!=columns[i].typeofcol)
                {
                    return false;
                }
            }
            return true;
        }

        public List<int> VerifyColumnValueIn(int index,List<string> list)
        {
            List<int> errores = new List<int>();
            var c = data.Columns[index];
            int i = 0;
            foreach (DataRow dtRow in data.Rows)
            {
                var pos = dtRow[c].ToString();
                if (!list.Exists(x => string.Equals(x, pos, StringComparison.OrdinalIgnoreCase)))
                {
                    errores.Add(i);
                }
                i++;
            }

            return errores;
        }

        public void paintColumnByIndex()
        {

        }
    }
}