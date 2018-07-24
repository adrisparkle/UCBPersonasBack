using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
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
using UcbBack.Logic.B1;
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
        public bool valid { get; set; }
        private int sheets { get; set; }
        public int headerin { get; private set; }
        private HanaValidator hanaValidator;
        //Image logo = Image.FromFile(HttpContext.Current.Server.MapPath("~/Images/logo.png"));
        private dynamic errors = new JObject();
        private ValidatePerson personValidator;
        private ApplicationDbContext _context;

        public ValidateExcelFile(Excelcol[] columns, string fileName, int headerin = 1,ApplicationDbContext context=null)
        {
            this.columns = columns;
            this.fileName = fileName;
            this.headerin = headerin;
            _context = context ?? new ApplicationDbContext();
            valid = true;
        }

        public ValidateExcelFile(Excelcol[] columns, Stream data, string fileName, int headerin =1, int sheets = 1, string resultfileName = "Result", ApplicationDbContext context=null)
        {
            _context = context ?? new ApplicationDbContext();
            this.columns = columns;
            this.fileName = fileName;
            this.resultfileName = resultfileName;
            this.sheets = sheets;
            this.headerin = headerin;
            this.wb = setExcelFile(data);
            hanaValidator = new HanaValidator();
            personValidator = new ValidatePerson();
            valid = true;
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
            IXLWorksheet ws =template.AddWorksheet(fileName.Replace(".xlsx",""));
            var tittle = ws.Range(1, 1, 2, columns.Length);
            tittle.Cell(1, 1).Value = fileName.Replace(".xlsx", "").ToUpper();//"Base de Datos Nacional de Capital Humano";
            tittle.Cell(1, 1).Style.Font.Bold = true;
            tittle.Cell(1, 1).Style.Fill.BackgroundColor = XLColor.FromTheme(XLThemeColor.Accent1);
            tittle.Cell(1, 1).Style.Font.FontName = "Bahnschrift SemiLight";
            tittle.Cell(1, 1).Style.Font.FontSize = 20;
            tittle.Cell(1, 1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            tittle.Merge();
            for (int i = 0; i < columns.Length; i++)
            {
                ws.Column(i + 1).Width=13;
                ws.Cell(headerin, i + 1).Value = columns[i].headers;
                /*if(columns[i].typeofcol==typeof(double))
                    for (int j = headerin + 1; j < 1000; j++)
                    {
                        ws.Cell(j, i + 1).Style.NumberFormat.Format = "#,##0.00";
                        var validation = ws.Cell(j, i + 1).DataValidation;
                        validation.Decimal.Between(0, 9999999);
                        validation.InputTitle = "Columna Numerica";
                        validation.InputMessage = "Por favor ingresar solo numeros.";
                        validation.ErrorStyle = XLErrorStyle.Warning;
                        validation.ErrorTitle = "Error de tipo de valor";
                        validation.ErrorMessage = "Esta celda debe ser tipo numerica.";
                    }*/

                ws.Cell(headerin, i + 1).Style.Alignment.WrapText = true;
                ws.Cell(headerin, i + 1).Style.Font.Bold = true;
                ws.Cell(headerin, i + 1).Style.Fill.BackgroundColor = XLColor.FromTheme(XLThemeColor.Accent1);
            }

            valid = true;
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

            for(int l = 1 ;l<=sheets;l++)
            {
                var sheet = wb.Worksheet(l);
                IXLRange UsedRange = sheet.RangeUsed();
                if (UsedRange.ColumnCount() != columns.Length)
                {
                    addError("Cantidad de Columnas", "Se esperaba tener " + columns.Length + "columnas en la hoja: " + sheet.Name + " se encontró " + UsedRange.ColumnCount());
                    res = false;
                }

                if (UsedRange.LastRow().RowNumber() <= headerin)
                {
                    addError("Archivo Sin Datos", "No se encontró datos en el archivo subido.");
                    res = false;
                }
                for (int i = 1; i <= columns.Length; i++)
                {
                    var comp = String.Compare(
                        Regex.Replace(sheet.Cell(headerin, i).Value.ToString().Trim().ToUpper(), @"\t|\n|\r", " "),
                        columns[i - 1].headers.Trim().ToUpper(), CultureInfo.CurrentCulture, CompareOptions.IgnoreNonSpace);
                    if (comp!=0)
                    {
                        res = false;
                        addError("Nombre de columna", "La columna " + i + "deberia llamarse: " + columns[i - 1].headers,false);
                        paintXY(i, headerin, XLColor.Red, "Esta Columna deberia llamarse: " + columns[i - 1].headers);
                    }
                    bool tipocol = true;
                    if (columns[i - 1].typeofcol != typeof(string))    
                    for (int j = headerin + 1; j < UsedRange.LastRow().RowNumber(); j++)
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

            valid = valid && res;
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
            string colId=null,
            bool notin=false)
        {
            bool res = true;
            IXLRange UsedRange = wb.Worksheet(sheet).RangeUsed();
            var l = UsedRange.LastRow().RowNumber();
            for (int i = headerin + 1; i <= UsedRange.LastRow().RowNumber(); i++)
            {
                if (list.Exists(x => string.Equals(x, wb.Worksheet(sheet).Cell(i, index).Value.ToString(), StringComparison.OrdinalIgnoreCase))==notin)
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

        
        public bool VerifyPerson(int ci = -1, int CUNI = -1, int fullname = -1, int sheet = 1, bool paintcolci = true, bool paintcolcuni = true, bool paintcolnombre = true, bool jaro = true, string comment = "No se encontro este valor en la Base de Datos Nacional.", bool personActive = true, string date = null, string format = "yyyy-MM-dd", int branchesId =-1)
        {
            bool res = true;
            IXLRange UsedRange = wb.Worksheet(sheet).RangeUsed();
            var c = new ApplicationDbContext();
            var ppllist = c.Person.ToList();

            for (int i = headerin + 1; i <= UsedRange.LastRow().RowNumber(); i++)
            {
                string strname = null;
                string strfsn = null;
                string strssn = null;
                string strmsn = null;
                try
                {
                    string strci = ci != -1 ? wb.Worksheet(sheet).Cell(i, ci).Value.ToString() : null;
                    string strcuni = CUNI != -1 ? wb.Worksheet(sheet).Cell(i, CUNI).Value.ToString() : null;

                    if (fullname != -1)
                    {
                        strfsn = wb.Worksheet(sheet).Cell(i, fullname).Value.ToString() == "" ? null : wb.Worksheet(sheet).Cell(i, fullname).Value.ToString();
                        strssn = wb.Worksheet(sheet).Cell(i, fullname + 1).Value.ToString() == "" ? null : wb.Worksheet(sheet).Cell(i, fullname + 1).Value.ToString();
                        strname = wb.Worksheet(sheet).Cell(i, fullname + 2).Value.ToString() == "" ? null : wb.Worksheet(sheet).Cell(i, fullname + 2).Value.ToString();
                        strmsn = wb.Worksheet(sheet).Cell(i, fullname + 3).Value.ToString() == "" ? null : wb.Worksheet(sheet).Cell(i, fullname + 3).Value.ToString();
                    }

                    if (!ppllist.Any(x => x.Document == strci
                                        && x.CUNI == strcuni
                                        && x.FirstSurName == strfsn
                                        && x.SecondSurName == strssn
                                        && x.Names == strname
                                        && x.MariedSurName == strmsn))
                    {
                        var p = ppllist.FirstOrDefault(x => x.Document == strci);
                        if (personActive && !personValidator.IsActive(p, date, format, branchId: branchesId))
                        {
                            res = false;
                            if (fullname != -1)
                            {
                                paintXY(fullname, i, XLColor.Red, "Esta Persona NO se encuentra Activa\n");
                                paintXY(fullname + 1, i, XLColor.Red, "Esta Persona NO se encuentra Activa\n");
                                paintXY(fullname + 2, i, XLColor.Red, "Esta Persona NO se encuentra Activa\n");
                                paintXY(fullname + 3, i, XLColor.Red, "Esta Persona NO se encuentra Activa\n");

                            }
                            if (ci != -1)
                                paintXY(ci, i, XLColor.Red, "Esta Persona NO se encuentra Activa\n");
                            if (CUNI != -1)
                                paintXY(CUNI, i, XLColor.Red, "Esta Persona NO se encuentra Activa\n");
                        }
                        if (strci != null && ppllist.Any(x => x.Document == strci.ToString()))
                        {
                            if (strname != null && strfsn != null && strssn != null && strmsn != null && !ppllist.Any(x => x.Document == strci
                                        && x.CUNI == strcuni
                                        && x.FirstSurName == strfsn
                                        && x.SecondSurName == strssn
                                        && x.Names == strname
                                        && x.MariedSurName == strmsn))
                            {
                                if (paintcolnombre)
                                {
                                    res = false;
                                    string aux = "";
                                    var similarities = ppllist.Where(x => x.Document == strci.ToString()).Select(y => new { y.FirstSurName, y.SecondSurName, y.Names, y.MariedSurName }).ToList();
                                    aux = similarities.Any() && strfsn != similarities[0].FirstSurName ? "\nNo será: '" + similarities[0].FirstSurName + "'?" : "";
                                    paintXY(fullname, i, XLColor.Red,  aux);

                                    aux = similarities.Any() && strssn != similarities[0].SecondSurName ? "\nNo será: '" + similarities[0].SecondSurName + "'?" : "";
                                    
                                    paintXY(fullname + 1, i, XLColor.Red, aux);

                                    aux = similarities.Any() && strname != similarities[0].Names ? "\nNo será: '" + similarities[0].Names + "'?" : "";
                                    paintXY(fullname + 2, i, XLColor.Red, aux);

                                    aux = similarities.Any() && strmsn != similarities[0].MariedSurName ? "\nNo será: '" + similarities[0].MariedSurName + "'?" : "";
                                    paintXY(fullname + 3, i, XLColor.Red, aux);
                                }
                            }
                            if (strcuni != null && !ppllist.Any(x => x.Document == wb.Worksheet(sheet).Cell(i, ci).Value.ToString()
                                               && x.CUNI == wb.Worksheet(sheet).Cell(i, CUNI).Value.ToString()))
                            {
                                if (paintcolcuni)
                                {
                                    res = false;
                                    string aux = "";
                                    var similarities = ppllist.Where(x => x.Document == strci).Select(y => y.CUNI).ToList();
                                    aux = similarities.Any() ? "\nNo será: '" + similarities[0].ToString() + "'?" : "";
                                    //wb.Worksheet(sheet).Cell(i, CUNI).Value = similarities;
                                    paintXY(CUNI, i, XLColor.Red, comment + aux);
                                }
                            }
                        }
                        else if (strcuni != null && ppllist.Any(x => x.CUNI == strcuni.ToString()))
                        {
                            if (strname != null && !ppllist.Any(x => x.CUNI == strcuni
                                                                     && x.FirstSurName == strfsn
                                                                     && x.SecondSurName == strssn
                                                                     && x.Names == strname
                                                                     && x.MariedSurName == strmsn))
                            {
                                if (paintcolnombre)
                                {
                                    res = false;
                                    string aux = "";
                                    var similarities = ppllist.Where(x => x.Document == strci).Select(y => new { y.FirstSurName, y.SecondSurName, y.Names, y.MariedSurName }).ToList();
                                    aux = similarities.Any() && strfsn != similarities[0].FirstSurName ? "\nNo será: '" + similarities[0].FirstSurName + "'?" : "";
                                    paintXY(fullname, i, XLColor.Red,  aux);

                                    aux = similarities.Any() && strssn != similarities[0].SecondSurName ? "\nNo será: '" + similarities[0].SecondSurName + "'?" : "";
                                    paintXY(fullname + 1, i, XLColor.Red,  aux);

                                    aux = similarities.Any() && strname != similarities[0].Names ? "\nNo será: '" + similarities[0].Names + "'?" : "";
                                    paintXY(fullname + 2, i, XLColor.Red,  aux);

                                    aux = similarities.Any() && strmsn != similarities[0].MariedSurName ? "\nNo será: '" + similarities[0].MariedSurName + "'?" : "";
                                    paintXY(fullname + 3, i, XLColor.Red,  aux);
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
                        else if (strname != null && ppllist.Any(x => x.FirstSurName == strfsn
                                                                 && x.SecondSurName == strssn
                                                                 && x.Names == strname
                                                                 /*&& x.MariedSurName == strmsn*/))
                        {
                            if (strcuni != null && !ppllist.Any(x => x.CUNI == strcuni
                                                  && x.SecondSurName == strssn
                                                  && x.Names == strname
                                                  && x.MariedSurName == strmsn))
                            {
                                if (paintcolcuni)
                                {
                                    res = false;
                                    string aux = "";
                                    var similarities = ppllist.Where(x => x.FirstSurName == strfsn
                                                                           && x.SecondSurName == strssn
                                                                           && x.Names == strname
                                                                           /*&& x.MariedSurName == strmsn*/).Select(y => y.CUNI).ToList();
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
                                    var similarities = ppllist.Where(x => x.FirstSurName == strfsn
                                                                          && x.SecondSurName == strssn
                                                                          && x.Names == strname
                                                                          && x.MariedSurName == strmsn).Select(y => y.Document).ToList();

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
                                var similarities = hanaValidator.Similarities(strfsn + " " + strssn + " " + strname, "'concat(a.\"FirstSurName\"," + "concat('' ''," + "concat(a.\"SecondSurName\"," + "concat('' '',a.\"Names\")" + ")" + ")" + ")'", "People", "\"CUNI\"", 0.9f);
                                if (similarities.Count > 0)
                                {
                                    string si = similarities[0];
                                    var person = _context.Person.FirstOrDefault(pe => pe.CUNI == si);
                                    

                                    aux = strfsn != person.FirstSurName ? "\nNo será: '" + person.FirstSurName + "'?" : "";
                                    paintXY(fullname, i, XLColor.Red,  aux);

                                    aux = strssn != person.SecondSurName ? "\nNo será: '" + person.SecondSurName + "'?" : "";
                                    paintXY(fullname + 1, i, XLColor.Red, aux);

                                    aux = strname != person.Names ? "\nNo será: '" + person.Names + "'?" : "";
                                    paintXY(fullname + 2, i, XLColor.Red,  aux);

                                    aux = strmsn != person.MariedSurName ? "\nNo será: '" + person.MariedSurName + "'?" : "";
                                    paintXY(fullname + 3, i, XLColor.Red,  aux);

                                    aux = strci != person.Document ? "\nNo será: '" + person.Document + "'?" : "";
                                    paintXY(ci, i, XLColor.Red,  aux);

                                    aux = strcuni != person.CUNI ? "\nNo será: '" + person.CUNI + "'?" : "";
                                    paintXY(CUNI, i, XLColor.Red, aux);
                                }
                                                                
                            }
                        }


                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                    throw;
                }
                
            }
            if(!res)
                addError("Datos Personas","Algunos datos de personas no coinciden o no existen en la Base de datos Nacional.");
            valid = valid && res;
            return res;

        }

        public bool VerifyParalel(int cod,int periodo, int sigla,int sheet =1)
        {
            var B1conn = B1Connection.Instance;
            bool res = true;
            IXLRange UsedRange = wb.Worksheet(sheet).RangeUsed();
            var c = new ApplicationDbContext();
            List<dynamic> list = B1conn.getParalels();

            for (int i = headerin + 1; i <= UsedRange.LastRow().RowNumber(); i++)
            {
                var strcod = cod != -1 ? wb.Worksheet(sheet).Cell(i, cod).Value.ToString() : null;
                var strperiodo = periodo != -1 ? wb.Worksheet(sheet).Cell(i, periodo).Value.ToString() : null;
                var strsigla = sigla != -1 ? wb.Worksheet(sheet).Cell(i, sigla).Value.ToString() : null;
                if (!list.Any(x => x.cod == strcod && x.periodo == strperiodo && x.sigla == strsigla))
                {
                    res = false;
                    if (list.Any(x => x.cod==strcod))
                    {
                        if (list.Any(x => x.cod == strcod && x.periodo == strperiodo))
                        {
                            paintXY(sigla, i, XLColor.Red, "Esta Sigla no es correcta." );
                        }
                        else if (list.Any(x => x.cod == strcod && x.sigla == strsigla))
                        {
                            paintXY(periodo, i, XLColor.Red, "Este Periodo no es correcto.");
                        }
                        else
                        {
                            paintXY(sigla, i, XLColor.Red, "Esta Sigla no es correcta.");
                            paintXY(periodo, i, XLColor.Red, "Este Periodo no es correcto.");
                        }
                    }
                    else if (list.Any(x => x.periodo==strperiodo))
                    {
                        if (list.Any(x => x.periodo == strperiodo && x.sigla == strsigla))
                        {
                            paintXY(cod, i, XLColor.Red, "Este Codigo no es correcto.");
                            //string co = list.FirstOrDefault(l => l.periodo == strperiodo && l.sigla == strsigla).cod;
                            //wb.Worksheet(1).Cell(i, cod).Value = co;
                        }
                        else
                        {
                            paintXY(cod, i, XLColor.Red, "Este Codigo no es correcto.");
                            paintXY(sigla, i, XLColor.Red, "Esta Sigla no es correcta.");
                        }
                    }
                    else if (list.Any(x => x.sigla==strsigla))
                    {
                        paintXY(cod, i, XLColor.Red, "Este Codigo no es correcto.");
                        paintXY(periodo, i, XLColor.Red, "Este Periodo no es correcto.");
                    }
                    else
                    {
                        paintXY(cod, i, XLColor.Red, "Este Codigo no es correcto.");
                        paintXY(periodo, i, XLColor.Red, "Este Periodo no es correcto.");
                        paintXY(sigla, i, XLColor.Red, "Este Periodo no es correcto.");
                    }
                }
            }

            valid = valid && res;
            return res;
        }

        public Decimal strToDecimal(int row, int col, int sheet=1)
        {
            return wb.Worksheet(sheet).Cell(row, col).Value.ToString() == "" ? 0.0m : Decimal.Parse(wb.Worksheet(sheet).Cell(row, col).Value.ToString());
        }
        public Double strToDouble(int row, int col, int sheet=1)
        {
           return wb.Worksheet(sheet).Cell(row, col).Value.ToString() == "" ? 0.0 : Double.Parse(wb.Worksheet(sheet).Cell(row, col).Value.ToString());            
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