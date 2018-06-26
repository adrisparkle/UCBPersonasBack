using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
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
using Newtonsoft.Json.Linq;
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
        public int headerin { get; private set; }
        private HanaValidator hanaValidator;
        //Image logo = Image.FromFile(HttpContext.Current.Server.MapPath("~/Images/logo.png"));
        private dynamic errors = new JObject();


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

        public HttpResponseMessage getTemplate()
        {
            var template = new XLWorkbook();
            IXLWorksheet ws =template.AddWorksheet(resultfileName);

            for (int i = 1; i <= columns.Length; i++)
            {
                template.Worksheet(1).Cell(3, i).Value = columns[i].headers;
            }
            return toResponse(template);
        }

        public bool isFormatValid()
        {
            bool res = true;
            if (sheets != wb.Worksheets.Count)
            {
                addError("Cantidad de Hojas", "Se envio un archivo con " + wb.Worksheets.Count + " hojas, se esperaba tener " + sheets + ", solo se revisó la" + (sheets > 1 ? "s" : "") + " primera" + (sheets > 1 ? "s" : "") + " hoja" + (sheets > 1 ? "s" : "") );
                res = false;
            }

            //foreach (IXLWorksheet sheet in wb.Worksheets && k--)
            for(int l = 1 ;l<=sheets;l++)
            {
                var sheet = wb.Worksheet(l);
                IXLRange UsedRange = sheet.RangeUsed();
                if (UsedRange.ColumnCount() != columns.Length)
                {
                    addError("Cantidad de Columnas", "Se esperaba tener " + columns.Length + "columnas en la hoja: " + sheet.Name + " se encontró " + UsedRange.ColumnCount());
                    res = false;
                }
                for (int i = 1; i <= columns.Length; i++)
                {
                    if (!String.Equals(Regex.Replace(sheet.Cell(headerin, i).Value.ToString().Trim(), @"\t|\n|\r", " "),
                        columns[i - 1].headers.Trim()))
                    {
                        res = false;
                        addError("Nombre de columna", "La columna " + i + "deberia llamarse: " + columns[i - 1].headers,false);
                        paintXY(i, headerin, XLColor.Red, "Esta Columna deberia llamarse: " + columns[i - 1].headers);
                    }
                    bool tipocol = true;
                    if (columns[i - 1].typeofcol != typeof(string))    
                    for (int j = headerin + 1; j < UsedRange.RowCount(); j++)
                    {
                        
                        if (sheet.Cell(j, i).Value.GetType() != columns[i - 1].typeofcol)
                        {
                            res = false;
                            var xx = sheet.Cell(j, i).Value;
                            paintXY(i, j, XLColor.Red, "Esta Celda deberia ser tipo: " + columns[i - 1].typeofcol.Name);
                            if (tipocol)
                            {
                                addError("Tipo de valor de columna", "La columna " + i + " deberia ser tipo: " + columns[i - 1].typeofcol.Name, false);
                                tipocol = false;
                            }
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
            var l = UsedRange.LastRow().RowNumber();
            for (int i = headerin + 1; i <= UsedRange.LastRow().RowNumber(); i++)
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
            if(!res)
                addError("Valor no valido", "Valor o valores no validos en la columna: "+index,false);
            return res;
        }

        //todo Verify if person is active
        public bool VerifyPerson(int ci=-1, int CUNI=-1, int fullname=-1, int sheet = 1, bool paintcolci = true, bool paintcolcuni = true, bool paintcolnombre = true, bool jaro = true,string comment ="No se encontro este valor en la Base de Datos Nacional.",bool personActive = true)
        {
            bool res = true;
            IXLRange UsedRange = wb.Worksheet(sheet).RangeUsed();
            var c = new ApplicationDbContext();
            var ppllist = c.Person.ToList();

            for (int i = headerin + 1; i < UsedRange.RowCount(); i++)
            {
                var strci = ci != -1 ? wb.Worksheet(sheet).Cell(i, ci).Value.ToString() : null;
                var strcuni = CUNI != -1 ? wb.Worksheet(sheet).Cell(i, CUNI).Value.ToString() : null;
                var strname = fullname != -1 ? wb.Worksheet(sheet).Cell(i, fullname).Value.ToString() : null;
                if (!ppllist.Any(x=>x.Document==strci
                                    && x.CUNI == strcuni
                                    && (x.FirstSurName + " " + x.SecondSurName + " " + x.Names) == strname))
                {
                    
                    if (strci!= null && ppllist.Any(x => x.Document == strci))
                    {
                        if (strname!=null && !ppllist.Any(x => x.Document == strci
                                              && (x.FirstSurName + " " + x.SecondSurName + " " + x.Names) ==
                                             strname))
                        {
                            if (paintcolnombre)
                            {
                                res = false;
                                string aux = "";
                                var similarities = ppllist.Where(x => x.Document == strci).Select(y => (y.FirstSurName + " " + y.SecondSurName + " " + y.Names)).ToList();
                                aux = similarities.Any() ? "\nNo será: '" + similarities[0].ToString() + "'?" : "";
                                paintXY(fullname, i, XLColor.Red, comment + aux);
                            }
                        }
                        if (strcuni != null && !ppllist.Any(x => x.Document == wb.Worksheet(sheet).Cell(i, ci).Value.ToString()
                                           && x.CUNI == wb.Worksheet(sheet).Cell(i, CUNI).Value.ToString()))
                        {
                            if (paintcolcuni)
                            {
                                res = false;
                                string aux = "";
                                var similarities = ppllist.Where(x => x.Document ==strci).Select(y => y.CUNI).ToList();
                                aux = similarities.Any() ? "\nNo será: '" + similarities[0].ToString() + "'?" : "";
                                paintXY(CUNI, i, XLColor.Red, comment + aux);
                            }
                        }
                    }
                    else if (strcuni != null && ppllist.Any(x => x.CUNI == strcuni))
                    {
                        if (strname != null && !ppllist.Any(x => x.CUNI == strcuni
                                              && (x.FirstSurName + " " + x.SecondSurName + " " + x.Names) ==
                                              strname))
                        {
                            if (paintcolnombre)
                            {
                                res = false;
                                string aux = "";
                                var similarities = ppllist.Where(x => x.CUNI == strcuni).Select(y => (y.FirstSurName + " " + y.SecondSurName + " " + y.Names)).ToList();
                                aux = similarities.Any() ? "\nNo será: '" + similarities[0].ToString() + "'?" : "";
                                paintXY(fullname, i, XLColor.Red, comment + aux);
                            }
                        }
                        if (strci != null && !ppllist.Any(x => x.Document == strci
                                              && x.CUNI == strcuni))
                        {
                            if (paintcolci)
                            {
                                res = false;
                                string aux = "";
                                var similarities = ppllist.Where(x => x.CUNI == strcuni).Select(y => y.Document).ToList();
                                aux = similarities.Any() ? "\nNo será: '" + similarities[0].ToString() + "'?" : "";
                                paintXY(ci, i, XLColor.Red, comment + aux);
                            }
                        }
                    }
                    else if (strname != null && ppllist.Any(x => (x.FirstSurName + " " + x.SecondSurName + " " + x.Names) ==
                                         strname))
                    {
                        if (strcuni != null && !ppllist.Any(x => x.CUNI == strcuni
                                              && (x.FirstSurName + " " + x.SecondSurName + " " + x.Names) ==
                                              strname))
                        {
                            if (paintcolcuni)
                            {
                                res = false;
                                string aux = "";
                                var similarities = ppllist.Where(x => (x.FirstSurName + " " + x.SecondSurName + " " + x.Names) ==
                                                                      strname).Select(y => y.CUNI).ToList();
                                aux = similarities.Any() ? "\nNo será: '" + similarities[0].ToString() + "'?" : "";
                                paintXY(CUNI, i, XLColor.Red, comment + aux);
                            }
                        }
                        if (strci != null && !ppllist.Any(x => x.Document == strci
                                              && (x.FirstSurName + " " + x.SecondSurName + " " + x.Names) ==
                                              strname))
                        {
                            if (paintcolci)
                            {
                                res = false;
                                string aux = "";
                                var similarities = ppllist.Where(x => (x.FirstSurName + " " + x.SecondSurName + " " + x.Names) ==
                                                                      strname).Select(y => y.Document).ToList();

                                aux = similarities.Any() ? "\nNo será: '" + similarities[0].ToString() + "'?" : "";
                                paintXY(ci, i, XLColor.Red, comment + aux);
                            }
                        }
                    }
                    else
                    {
                        string aux = "";
                        if (strname != null && jaro)
                        {
                            var similarities = hanaValidator.Similarities(strname, "'concat(a.\"FirstSurName\"," +"concat('' ''," +"concat(a.\"SecondSurName\"," +"concat('' '',a.\"Names\")" +")" +")" +")'", "People", "concat(a.\"FirstSurName\",concat('' '',concat(a.\"SecondSurName\",concat('' '',a.\"Names\"))))", 0.9f);
                            aux = similarities.Any() ? "\nNo será: '" + similarities[0].ToString() + "'?" : "";
                        }
                      
                        if (paintcolnombre)
                        {
                            res = false;
                            paintXY(fullname, i, XLColor.Red, comment + aux);
                        }
                    }


                }
            }
            if(!res)
                addError("Datos Personas","Algunos datos de personas no coinciden o no existen en la Base de datos Nacional.");
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

        public HttpResponseMessage toResponse(XLWorkbook w = null)
        {
            w = w ?? wb;
            HttpResponseMessage response = new HttpResponseMessage();
            var ms = new MemoryStream();
            if (w != null)
            {
                w.SaveAs(ms);
                response.StatusCode = valid ? HttpStatusCode.OK : HttpStatusCode.BadRequest;
                response.Content = new StreamContent(ms);
                response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");
                response.Content.Headers.ContentDisposition.FileName = resultfileName + ".xlsx";
                response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
                response.Content.Headers.ContentLength = ms.Length;
                response.Headers.Add("UploadErrors",errors.ToString().Replace("\r\n",""));
                ms.Seek(0, SeekOrigin.Begin); 
            }
            else
            {
                response.StatusCode = HttpStatusCode.BadRequest;
                response.Headers.Add("UploadErrors", errors.ToString());
                response.Content = new StringContent("Formato del archivo no valido.");
            }
            return response;
        }

        public XLWorkbook setExcelFile(Stream stream)
        {
            wb = new XLWorkbook(stream);
            return wb;
        }

        public void addError(string error_name,string err,bool replace = true)
        {
            error_name = HttpUtility.HtmlEncode(error_name);
            err = HttpUtility.HtmlEncode(err);
            errors[error_name] = replace ? err : errors[error_name] == null ? errors[error_name] = err : errors[error_name]+","+err;
        }
    }
}