using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.RegularExpressions;
using System.Web;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2013.Word;
using ExcelDataReader;
using UcbBack.Models;

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
    public abstract class ValidateExcelFile
    {
        private Excelcol[] columns { get; set; }
        //private DataTable data { get; set; }
        private string fileName { get; set; }
        private string resultfileName { get; set; }
        public XLWorkbook wb { get; private set; }
        private bool valid { get; set; }
        private int sheets { get; set; }
        private int headerin { get; set; }
        private HanaValidator hanaValidator;

        public ValidateExcelFile(Excelcol[] columns, Stream data, string fileName, int headerin =1, int sheets = 1, string resultfileName = "Result")
        {
            this.columns = columns;
            this.fileName = fileName;
            this.resultfileName = resultfileName;
            this.sheets = sheets;
            this.headerin = headerin;
            //data = setExcelFile(d); 
            this.wb = setExcelFile(data);
            hanaValidator = new HanaValidator();
        }

        public abstract void toDataBase();
        public abstract bool ValidateFile();

        public bool isValid()
        {
            return valid;
        }

        public bool isFormatValid()
        {
            bool res = true;
            if (sheets != wb.Worksheets.Count)
                return false;
            
            foreach (IXLWorksheet sheet in wb.Worksheets)
            {
                IXLRange UsedRange = sheet.RangeUsed();
                if (UsedRange.ColumnCount() != columns.Length)
                    return false;
                for (int i = 1; i <= columns.Length; i++)
                {
                    if (!String.Equals(Regex.Replace(sheet.Cell(headerin, i).Value.ToString().Trim(), @"\t|\n|\r", " "),
                        columns[i - 1].headers.Trim()))
                    {
                        res = false;
                        paintXY(i, headerin, XLColor.Red, "Esta Columna deberia llamarse: " + columns[i - 1].headers);
                    }
                    if (columns[i - 1].typeofcol != typeof(string))    
                    for (int j = headerin + 1; j < UsedRange.RowCount() + headerin; j++)
                    {
                        if (sheet.Cell(j, i).Value.GetType() != columns[i - 1].typeofcol)
                        {
                            res = false;
                            paintXY(i, j, XLColor.Red, "Esta Celda deberia ser tipo: " + columns[i - 1].typeofcol.Name);
                        }
                    }
                }
            }

            valid = res;
            return res;
        }

        public bool VerifyColumnValueIn(
            int index,
            List<string> list,
            bool paintcol=true,
            int sheet=1,
            string comment = "Este Valor no es permitido en esta columna.",
            bool jaro=false,
            string colToCompare=null, 
            string table=null, 
            string colId=null)
        {
            bool res = true;
            IXLRange UsedRange = wb.Worksheet(sheet).RangeUsed();

            for (int i = headerin + 1; i <= UsedRange.RowCount()+headerin; i++)
            {
                if (!list.Exists(x => string.Equals(x, wb.Worksheet(sheet).Cell(i, index).Value.ToString(), StringComparison.OrdinalIgnoreCase)))
                {
                    res = false;
                    if (paintcol)
                    {
                        string aux = "";
                        if (jaro)
                        {
                            var similarities = hanaValidator.Similarities(wb.Worksheet(sheet).Cell(i, index).Value.ToString(), colToCompare, table, colId, 0.9f);

                            aux = similarities.Any() ? "\nNo será: '"+similarities[0].ToString()+"'?":"";
                        }
                        paintXY(index, i, XLColor.Red, comment + aux);
                    }
                        
                }
            }
            valid = valid && res;
            return res;
        }

        public bool VerifyPerson(int ci, int CUNI, int fullname, int sheet = 1, bool paintcolci = true, bool paintcolcuni = true, bool paintcolnombre = true, bool jaro = true,string comment ="No se encontro este valor en la Base de Datos Nacional.")
        {
            bool res = true;
            IXLRange UsedRange = wb.Worksheet(sheet).RangeUsed();
            var c = new ApplicationDbContext();
            var ppllist = c.Person.ToList();

            for (int i = headerin + 1; i < UsedRange.RowCount() + headerin; i++)
            {
                if (!ppllist.Any(x=>x.Document==wb.Worksheet(sheet).Cell(i, ci).Value.ToString()
                                    && x.CUNI == wb.Worksheet(sheet).Cell(i, CUNI).Value.ToString()
                                    && (x.FirstSurName + " " + x.SecondSurName + " " + x.Names) == wb.Worksheet(sheet).Cell(i, fullname).Value.ToString()))
                {
                    res = false;
                    if (ppllist.Any(x => x.Document == wb.Worksheet(sheet).Cell(i, ci).Value.ToString()))
                    {
                        if (!ppllist.Any(x => x.Document == wb.Worksheet(sheet).Cell(i, ci).Value.ToString()
                                              && (x.FirstSurName + " " + x.SecondSurName + " " + x.Names) ==
                                              wb.Worksheet(sheet).Cell(i, fullname).Value.ToString()))
                        {
                            if (paintcolnombre)
                            {
                                string aux = "";
                                var similarities = ppllist.Where(x => x.Document == wb.Worksheet(sheet).Cell(i, ci).Value.ToString()).Select(y => (y.FirstSurName + " " + y.SecondSurName + " " + y.Names)).ToList();
                                aux = similarities.Any() ? "\nNo será: '" + similarities[0].ToString() + "'?" : "";
                                paintXY(fullname, i, XLColor.Red, comment + aux);
                            }
                        }
                        if(!ppllist.Any(x=>x.Document==wb.Worksheet(sheet).Cell(i, ci).Value.ToString()
                                           && x.CUNI == wb.Worksheet(sheet).Cell(i, CUNI).Value.ToString()))
                        {
                            if (paintcolcuni)
                            {
                                string aux = "";
                                var similarities = ppllist.Where(x => x.Document == wb.Worksheet(sheet).Cell(i, ci).Value.ToString()).Select(y => y.CUNI).ToList();
                                aux = similarities.Any() ? "\nNo será: '" + similarities[0].ToString() + "'?" : "";
                                paintXY(CUNI, i, XLColor.Red, comment + aux);
                            }
                        }
                    }
                    else if (ppllist.Any(x => x.CUNI == wb.Worksheet(sheet).Cell(i, CUNI).Value.ToString()))
                    {
                        if (!ppllist.Any(x => x.CUNI == wb.Worksheet(sheet).Cell(i, CUNI).Value.ToString()
                                              && (x.FirstSurName + " " + x.SecondSurName + " " + x.Names) ==
                                              wb.Worksheet(sheet).Cell(i, fullname).Value.ToString()))
                        {
                            if (paintcolnombre)
                            {
                                string aux = "";
                                var similarities = ppllist.Where(x => x.CUNI == wb.Worksheet(sheet).Cell(i, CUNI).Value.ToString()).Select(y => (y.FirstSurName + " " + y.SecondSurName + " " + y.Names)).ToList();
                                aux = similarities.Any() ? "\nNo será: '" + similarities[0].ToString() + "'?" : "";
                                paintXY(fullname, i, XLColor.Red, comment + aux);
                            }
                        }
                        if (!ppllist.Any(x => x.Document == wb.Worksheet(sheet).Cell(i, ci).Value.ToString()
                                              && x.CUNI == wb.Worksheet(sheet).Cell(i, CUNI).Value.ToString()))
                        {
                            if (paintcolci)
                            {
                                string aux = "";
                                var similarities = ppllist.Where(x => x.CUNI == wb.Worksheet(sheet).Cell(i, CUNI).Value.ToString()).Select(y => y.Document).ToList();
                                aux = similarities.Any() ? "\nNo será: '" + similarities[0].ToString() + "'?" : "";
                                paintXY(ci, i, XLColor.Red, comment + aux);
                            }
                        }
                    }
                    else if (ppllist.Any(x => (x.FirstSurName + " " + x.SecondSurName + " " + x.Names) ==
                                         wb.Worksheet(sheet).Cell(i, fullname).Value.ToString()))
                    {
                        if (!ppllist.Any(x => x.CUNI == wb.Worksheet(sheet).Cell(i, CUNI).Value.ToString()
                                              && (x.FirstSurName + " " + x.SecondSurName + " " + x.Names) ==
                                              wb.Worksheet(sheet).Cell(i, fullname).Value.ToString()))
                        {
                            if (paintcolcuni)
                            {
                                string aux = "";
                                var similarities = ppllist.Where(x => (x.FirstSurName + " " + x.SecondSurName + " " + x.Names) ==
                                                                      wb.Worksheet(sheet).Cell(i, fullname).Value.ToString()).Select(y => y.CUNI).ToList();
                                aux = similarities.Any() ? "\nNo será: '" + similarities[0].ToString() + "'?" : "";
                                paintXY(CUNI, i, XLColor.Red, comment + aux);
                            }
                        }
                        if (!ppllist.Any(x => x.Document == wb.Worksheet(sheet).Cell(i, ci).Value.ToString()
                                              && (x.FirstSurName + " " + x.SecondSurName + " " + x.Names) ==
                                              wb.Worksheet(sheet).Cell(i, fullname).Value.ToString()))
                        {
                            if (paintcolci)
                            {
                                string aux = "";
                                var similarities = ppllist.Where(x => (x.FirstSurName + " " + x.SecondSurName + " " + x.Names) ==
                                                                      wb.Worksheet(sheet).Cell(i, fullname).Value.ToString()).Select(y => y.Document).ToList();

                                aux = similarities.Any() ? "\nNo será: '" + similarities[0].ToString() + "'?" : "";
                                paintXY(ci, i, XLColor.Red, comment + aux);
                            }
                        }
                    }
                    else
                    {
                        string aux = "";
                        if (jaro)
                        {
                            var similarities = hanaValidator.Similarities(wb.Worksheet(sheet).Cell(i, fullname).Value.ToString(), "'concat(a.\"FirstSurName\"," +"concat('' ''," +"concat(a.\"SecondSurName\"," +"concat('' '',a.\"Names\")" +")" +")" +")'", "People", "concat(a.\"FirstSurName\",concat('' '',concat(a.\"SecondSurName\",concat('' '',a.\"Names\"))))", 0.9f);
                            aux = similarities.Any() ? "\nNo será: '" + similarities[0].ToString() + "'?" : "";

                        }
                      
                        if (paintcolnombre)
                        {
                            paintXY(fullname, i, XLColor.Red, comment + aux);
                        }
                    }


                }
            }
            valid = valid && res;
            return res;

        }

        public void paintXY(int x, int y,XLColor color,string comment = null)
        {
            wb.Worksheet(1).Cell(y, x).Style.Fill.BackgroundColor = color;
            if (comment != null)
            {
                wb.Worksheet(1).Cell(y, x).Comment.Style.Alignment.SetAutomaticSize();
                wb.Worksheet(1).Cell(y, x).Comment.AddText(comment);
            }
        }

        public HttpResponseMessage toResponse()
        {
            HttpResponseMessage response = new HttpResponseMessage();
            var ms = new MemoryStream();
            if (wb != null)
            {
                wb.SaveAs(ms);
                response.StatusCode = valid ? HttpStatusCode.OK : HttpStatusCode.BadRequest;
                response.Content = new StreamContent(ms);
                response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");
                response.Content.Headers.ContentDisposition.FileName = resultfileName + ".xlsx";
                response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
                response.Content.Headers.ContentLength = ms.Length;
                ms.Seek(0, SeekOrigin.Begin); 
            }
            else
            {
                response.StatusCode = HttpStatusCode.BadRequest;
                response.Content = new StringContent("Formato del archivo no valido.");
            }
            return response;
        }

        public XLWorkbook setExcelFile(Stream stream)
        {
            var wb = new XLWorkbook(stream);
            return wb;
        }
    }
}