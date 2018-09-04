using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Formatting;
using System.Reflection;
using System.Threading.Tasks;
using System.Web;
using System.Web.Http;
using DocumentFormat.OpenXml.Packaging;
using Newtonsoft.Json.Linq;
using UcbBack.Logic.ExcelFiles;
using UcbBack.Models;
using System.Linq;
using System.Net.Http.Headers;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using UcbBack.Logic;
using UcbBack.Models.Not_Mapped;

namespace UcbBack.Controllers
{
    public class PayrollController : ApiController
    {
        private ApplicationDbContext _context;

        private struct ProcessState
        {
            public static string STARTED = "STARTED";
            public static string ERROR = "ERROR";
            public static string CANCELED = "CANCELED";
            public static string VALIDATED = "VALIDATED";
            public static string PROCESSED = "PROCESSED";
            public static string WARNING = "WARNING";
            public static string INSAP = "INSAP";
        }
        private struct FileState
        {
            public static string SENDED = "SENDED";
            public static string UPLOADED = "UPLOADED";
            public static string ERROR = "ERROR";
            public static string CANCELED = "CANCELED";
        }
        private enum ExcelFileType
        {
            Payroll = 1,
            Academic,
            Discount,
            Postgrado,
            Pregrado,
            OR
        }

        private int ExcelHeaders = 3;

        private ValidateAuth auth;
        public PayrollController()
        {
            auth = new ValidateAuth();
            _context = new ApplicationDbContext();
        }

        [HttpPost]
        [Route("api/payroll/CheckUpload")]
        public IHttpActionResult CheckUpload([FromBody] JObject upload)
        {
            int branchid = 0;
            if (upload["mes"] == null || upload["gestion"] == null || upload["segmentoOrigen"] == null || !Int32.TryParse(upload["segmentoOrigen"].ToString(), out branchid))
                return BadRequest("Debes enviar mes,gestion y segmentoOrigen");
            var mes = upload["mes"].ToString();
            var gestion = upload["gestion"].ToString();

            var process = _context.DistProcesses.FirstOrDefault(f => f.mes == mes
                                                             && f.gestion == gestion
                                                             && f.BranchesId == branchid
                                                             && f.State!=ProcessState.CANCELED);
            if (process == null)
                return Ok();
            var files = _context.FileDbs.Where(f => f.DistProcessId == process.Id && f.State == FileState.UPLOADED).Include(f => f.DistFileTypeId).Select(f => new { f.DistFileType.FileType });

            List<string> tipos = new List<string>();
            foreach (var tipo in files)
            {
                tipos.Add(tipo.FileType);
            }

            dynamic res = new JObject();
            res.array = JToken.FromObject(tipos);
            res.id = process.Id;
            res.state = process.State;
            return Ok(res);
        }

        [NonAction]
        private Dist_File AddFileToProcess(string mes, string gestion, int BranchesId,ExcelFileType FileType,int userid,string fileName)
        {
            var processInDB = _context.DistProcesses.FirstOrDefault(p =>
                    p.BranchesId == BranchesId && p.gestion == gestion && p.mes == mes && (p.State == ProcessState.STARTED || p.State == ProcessState.ERROR || p.State == ProcessState.WARNING));


            if (processInDB != null)
            {
                var fileInDB = _context.FileDbs.FirstOrDefault(f => f.DistProcessId == processInDB.Id && f.DistFileTypeId == (int)FileType && f.State == FileState.UPLOADED);
                if (fileInDB == null)
                {
                    processInDB.State = ProcessState.STARTED;
                    _context.Database.ExecuteSqlCommand("UPDATE ADMNALRRHH.\"Dist_LogErrores\" set \"Inspected\" = true where \"DistProcessId\" = " + processInDB.Id);
                    var file = new Dist_File();
                    file.Id = _context.Database.SqlQuery<int>("SELECT ADMNALRRHH.\"rrhh_Dist_File_sqs\".nextval FROM DUMMY;").ToList()[0];
                    file.UploadedDate = DateTime.Now;
                    file.DistFileTypeId = (int)FileType; 
                    file.Name = fileName;
                    file.State = FileState.SENDED;
                    file.CustomUserId = userid;
                    file.DistProcessId = processInDB.Id;
                    _context.FileDbs.Add(file);
                    _context.SaveChanges();
                    return file;
                }
            }
            else
            {
                var process = new Dist_Process();
                process.UploadedDate = DateTime.Now;
                process.Id = _context.Database.SqlQuery<int>("SELECT ADMNALRRHH.\"rrhh_Dist_Process_sqs\".nextval FROM DUMMY;").ToList()[0];
                process.BranchesId = BranchesId;
                process.mes = mes;
                process.gestion = gestion;
                process.State = ProcessState.STARTED;
                _context.DistProcesses.Add(process);
                _context.SaveChanges();

                var file = new Dist_File();
                file.Id = _context.Database.SqlQuery<int>("SELECT ADMNALRRHH.\"rrhh_Dist_File_sqs\".nextval FROM DUMMY;").ToList()[0];
                file.UploadedDate = DateTime.Now;
                file.DistFileTypeId = (int)FileType;
                file.Name = fileName;
                file.State = FileState.SENDED;
                file.CustomUserId = userid;
                file.DistProcessId = process.Id;
                _context.FileDbs.Add(file);
                _context.SaveChanges();
                return file;
            }

            return null;
        }

        [NonAction]
        private async Task<System.Dynamic.ExpandoObject> HttpContentToVariables(MultipartMemoryStreamProvider req)
        {
            dynamic res = new System.Dynamic.ExpandoObject();      
            foreach (HttpContent contentPart in req.Contents)
            {
                var contentDisposition = contentPart.Headers.ContentDisposition;
                string varname = contentDisposition.Name;
                if (varname == "\"mes\"")
                {
                    if(contentPart.ReadAsStringAsync().Result.Length==2)
                        res.mes = contentPart.ReadAsStringAsync().Result.ToString();
                }
                else if (varname == "\"gestion\"")
                {
                    if (contentPart.ReadAsStringAsync().Result.Length == 4)
                        res.gestion = contentPart.ReadAsStringAsync().Result.ToString();
                }
                else if (varname == "\"segmentoOrigen\"")
                {
                    res.segmentoOrigen = Int32.Parse(contentPart.ReadAsStringAsync().Result.ToString());
                }
                else if (varname == "\"file\"")
                {
                    Stream stream = await contentPart.ReadAsStreamAsync();
                    res.fileName = String.IsNullOrEmpty(contentDisposition.FileName) ? "" : contentDisposition.FileName.Trim('"');
                    res.excelStream = stream;
                }
            }
            return res;
        }

        [HttpGet]
        [Route("api/payroll/PayrollExcel")]
        public HttpResponseMessage GetPayrollExcel()
        {
            PayrollExcel contractExcel = new PayrollExcel(fileName: "Planilla.xlsx", headerin: ExcelHeaders);
            return contractExcel.getTemplate();
        }
        [HttpDelete]
        [Route("api/payroll/PayrollExcel")]
        public IHttpActionResult CancelPayrollExcel(JObject data)
        {
            int branchesid;
            if (data["mes"] == null || data["gestion"] == null || data["segmentoOrigen"] == null || !Int32.TryParse(data["segmentoOrigen"].ToString(),out branchesid))
            {
                ModelState.AddModelError("Mal Formato","Debes enviar mes, gestion y segmentoOrigen");
                return BadRequest();

            }
            string mes = data["mes"].ToString();
            string gestion = data["gestion"].ToString();

            var file = _context.FileDbs.Include(f => f.DistProcess)
                .FirstOrDefault(f => f.DistProcess.mes == mes
                                     && f.DistProcess.gestion == gestion
                                     && f.DistProcess.BranchesId == branchesid
                                     && f.DistFileTypeId == (int)ExcelFileType.Payroll
                                     && f.State == FileState.UPLOADED);
            if (file == null)
            {
                return NotFound();
            }

            file.State = FileState.CANCELED;
            _context.SaveChanges();

            return Ok();
        }

        [HttpPost]
        [Route("api/payroll/PayrollExcel")]
        public async Task<HttpResponseMessage> UploadPayrollExcel()
        {
            var response = new HttpResponseMessage();
            PayrollExcel ExcelFile = null;
            try
            {
                var req = await Request.Content.ReadAsMultipartAsync();
                dynamic o = HttpContentToVariables(req).Result;

                if (!((IDictionary<string, object>)o).ContainsKey("mes")
                    || !((IDictionary<string, object>)o).ContainsKey("gestion")
                    || !((IDictionary<string, object>)o).ContainsKey("segmentoOrigen")
                    || !((IDictionary<string, object>)o).ContainsKey("fileName")
                    || !((IDictionary<string, object>)o).ContainsKey("excelStream")
                    || !o.fileName.ToString().EndsWith(".xlsx"))
                {
                    response.StatusCode = HttpStatusCode.BadRequest;
                    response.Headers.Add("UploadErrors", "{ \"Faltan datos\": \"Debe enviar mes(mm), gestion(yyyy), segmentoOrigen(id) y un archivo excel llamado file (en formato .xlsx)\"}");
                    response.Content = new StringContent("Debe enviar mes(mm), gestion(yyyy), segmentoOrigen(id) y un archivo excel llamado file");
                    return response;
                }

                int userid = 0;
                if (!Int32.TryParse(Request.Headers.GetValues("id").First(),out userid))
                {
                    response.StatusCode = HttpStatusCode.Unauthorized;
                    return response;
                }
                var file = AddFileToProcess(o.mes.ToString(), o.gestion.ToString(), o.segmentoOrigen, ExcelFileType.Payroll, userid, o.fileName.ToString());

                if (file == null)
                {
                    response.StatusCode = HttpStatusCode.BadRequest;
                    response.Headers.Add("UploadErrors", "{ \"Ya se Subio archivos para este mes\": \"Ya se subio  datos para este mes, si quiere volver a subir cancele el anterior archivo.\"}");
                    response.Content = new StringContent("Ya se subió  datos para este mes, si quiere volver a subir cancele el anterior archivo.");
                    return response;
                }

                ExcelFile = new PayrollExcel(o.excelStream, _context, o.fileName, o.mes, o.gestion, o.segmentoOrigen.ToString(), file, headerin: ExcelHeaders, sheets: 1);
                
                if (ExcelFile.ValidateFile())
                {
                    ExcelFile.toDataBase();
                    file.State = FileState.UPLOADED;
                    _context.SaveChanges();
                    response.StatusCode = HttpStatusCode.OK;
                    response.Content = new StringContent("Se subio el archivo correctamente.");
                    return response;
                }
                file.State = FileState.ERROR;
                _context.SaveChanges();
                return ExcelFile.toResponse();
            }
            catch (System.ArgumentException e)
            {
                response.StatusCode = HttpStatusCode.BadRequest;
                response.Headers.Add("UploadErrors", "{ \"Formato Archivo Invalido\": \"Por favor enviar un archivo en formato excel (.xlsx)\"}");
                ExcelFile.addError("Formato Archivo Invalido", "Por favor enviar un archivo en formato excel (.xlsx)");
                response.Content = new StringContent("Por favor enviar un archivo en formato excel (.xlsx)" + e);
                return response;
            }
            catch (System.IO.IOException e)
            {
                response.StatusCode = HttpStatusCode.BadRequest;
                response.Headers.Add("UploadErrors", "{ \"Archivo demasiado grande\": \"El archivo es demasiado grande para ser procesado.\"}");
                ExcelFile.addError("Archivo demasiado grande", "El archivo es demasiado grande para ser procesado.");
                response.Content = new StringContent("El archivo es demasiado grande para ser procesado.");
                return response;
            }
            catch (System.Exception e)
            {
                response.StatusCode = HttpStatusCode.BadRequest;
                response.Headers.Add("UploadErrors", "{ \"Existen Enlaces a otros archivos\": \"Existen celdas con referencias a otros archivos.\"}");
                response.Content = new StringContent("Por favor enviar un archivo en formato excel sin referencias a otros libros excel o formulas(.xls, .xslx)");
                return response;
            }
        }


        [HttpGet]
        [Route("api/payroll/AcademicExcel")]
        public HttpResponseMessage GetAcademicExcel()
        {
            AcademicExcel contractExcel = new AcademicExcel(fileName: "Academico.xlsx", headerin: ExcelHeaders);
            return contractExcel.getTemplate();
        }

        [HttpDelete]
        [Route("api/payroll/AcademicExcel")]
        public IHttpActionResult CancelAcademicExcel(JObject data)
        {
            int branchesid;
            if (data["mes"] == null || data["gestion"] == null || data["segmentoOrigen"] == null || !Int32.TryParse(data["segmentoOrigen"].ToString(), out branchesid))
            {
                ModelState.AddModelError("Mal Formato", "Debes enviar mes, gestion y segmentoOrigen");
                return BadRequest();

            }

            string mes = data["mes"].ToString();
            string gestion = data["gestion"].ToString();

            var file = _context.FileDbs.Include(f => f.DistProcess)
                .FirstOrDefault(f => f.DistProcess.mes == mes
                                     && f.DistProcess.gestion == gestion
                                     && f.DistProcess.BranchesId == branchesid
                                     && f.DistFileTypeId == (int) ExcelFileType.Academic
                                     && f.State == FileState.UPLOADED);
            if (file == null)
            {
                return NotFound();
            }

            file.State = FileState.CANCELED;
            _context.SaveChanges();

            return Ok();
        }

        [HttpPost]
        [Route("api/payroll/AcademicExcel")]
        public async Task<HttpResponseMessage> UploadAcademicExcel()
        {
            var response = new HttpResponseMessage();
            AcademicExcel ExcelFile=null;
            try
            {
                var req = await Request.Content.ReadAsMultipartAsync();
                dynamic o = HttpContentToVariables(req).Result;

                if (!((IDictionary<string, object>)o).ContainsKey("mes")
                    || !((IDictionary<string, object>)o).ContainsKey("gestion")
                    || !((IDictionary<string, object>)o).ContainsKey("segmentoOrigen")
                    || !((IDictionary<string, object>)o).ContainsKey("fileName")
                    || !((IDictionary<string, object>)o).ContainsKey("excelStream")
                    || !o.fileName.ToString().EndsWith(".xlsx"))
                {
                    response.StatusCode = HttpStatusCode.BadRequest;
                    response.Headers.Add("UploadErrors", "{ \"Faltan datos\": \"Debe enviar mes(mm), gestion(yyyy), segmentoOrigen(id) y un archivo excel llamado file (en formato .xlsx)\"}");
                    response.Content = new StringContent("Debe enviar mes(mm), gestion(yyyy), segmentoOrigen(id) y un archivo excel llamado file");
                    return response;
                }

                int userid = Int32.Parse(Request.Headers.GetValues("id").First());
                var file = AddFileToProcess(o.mes.ToString(), o.gestion.ToString(), o.segmentoOrigen, ExcelFileType.Academic, userid, o.fileName.ToString());

                if (file == null)
                {
                    response.StatusCode = HttpStatusCode.BadRequest;
                    response.Headers.Add("UploadErrors", "{ \"Ya se Subio archivos para este mes\": \"Ya se subio  datos para este mes, si quiere volver a subir cancele el anterior archivo.\"}");
                    response.Content = new StringContent("Ya se subió  datos para este mes, si quiere volver a subir cancele el anterior archivo.");
                    return response;
                }

                ExcelFile = new AcademicExcel(o.excelStream, _context, o.fileName, o.mes, o.gestion, o.segmentoOrigen.ToString(), file, headerin: ExcelHeaders, sheets: 1);

                if (ExcelFile.ValidateFile())
                {
                    ExcelFile.toDataBase();
                    response.StatusCode = HttpStatusCode.OK;
                    file.State = FileState.UPLOADED;
                    _context.SaveChanges();
                    response.Content = new StringContent("Se subio el archivo correctamente.");
                    return response;
                }
                file.State = FileState.ERROR;
                _context.SaveChanges();
                return ExcelFile.toResponse();
            }
            catch (System.ArgumentException e)
            {
                response.StatusCode = HttpStatusCode.BadRequest;
                response.Headers.Add("UploadErrors", "{ \"Formato Archivo Invalido\": \"Por favor enviar un archivo en formato excel (.xlsx)\"}");
                ExcelFile.addError("Formato Archivo Invalido", "Por favor enviar un archivo en formato excel (.xlsx)");
                response.Content = new StringContent("Por favor enviar un archivo en formato excel (.xlsx)" + e);
                return response;
            }
            catch (System.IO.IOException e)
            {
                response.StatusCode = HttpStatusCode.BadRequest;
                response.Headers.Add("UploadErrors", "{ \"Archivo demasiado grande\": \"El archivo es demasiado grande para ser procesado.\"}");
                ExcelFile.addError("Archivo demasiado grande", "El archivo es demasiado grande para ser procesado.");
                response.Content = new StringContent("El archivo es demasiado grande para ser procesado.");
                return response;
            }
            catch (System.Exception e)
            {
                response.StatusCode = HttpStatusCode.BadRequest;
                response.Headers.Add("UploadErrors", "{ \"Existen Enlaces a otros archivos\": \"Existen celdas con referencias a otros archivos.\"}");
                response.Content = new StringContent("Por favor enviar un archivo en formato excel sin referencias a otros libros excel o formulas(.xls, .xslx)");
                return response;
            }
        }

        [HttpGet]
        [Route("api/payroll/DiscountExcel")]
        public HttpResponseMessage GetDiscountExcel()
        {
            DiscountExcel contractExcel = new DiscountExcel(fileName: "Descuentos.xlsx", headerin: ExcelHeaders);
            return contractExcel.getTemplate();
        }

        [HttpDelete]
        [Route("api/payroll/DiscountExcel")]
        public IHttpActionResult CancelDiscountExcel(JObject data)
        {
            int branchesid;
            if (data["mes"] == null || data["gestion"] == null || data["segmentoOrigen"] == null || !Int32.TryParse(data["segmentoOrigen"].ToString(), out branchesid))
            {
                ModelState.AddModelError("Mal Formato", "Debes enviar mes, gestion y segmentoOrigen");
                return BadRequest();

            }
            string mes = data["mes"].ToString();
            string gestion = data["gestion"].ToString();
            var file = _context.FileDbs.Include(f => f.DistProcess)
                .FirstOrDefault(f => f.DistProcess.mes == mes
                                     && f.DistProcess.gestion == gestion
                                     && f.DistProcess.BranchesId == branchesid
                                     && f.DistFileTypeId == (int)ExcelFileType.Discount
                                     && f.State == FileState.UPLOADED);
            if (file == null)
            {
                return NotFound();
            }

            file.State = FileState.CANCELED;
            _context.SaveChanges();

            return Ok();
        }

        [HttpPost]
        [Route("api/payroll/DiscountExcel")]
        public async Task<HttpResponseMessage> UploadDiscountExcel()
        {
            var response = new HttpResponseMessage();
            DiscountExcel ExcelFile = null;
            try
            {
                var req = await Request.Content.ReadAsMultipartAsync();
                dynamic o = HttpContentToVariables(req).Result;

                if (!((IDictionary<string, object>)o).ContainsKey("mes")
                    || !((IDictionary<string, object>)o).ContainsKey("gestion")
                    || !((IDictionary<string, object>)o).ContainsKey("segmentoOrigen")
                    || !((IDictionary<string, object>)o).ContainsKey("fileName")
                    || !((IDictionary<string, object>)o).ContainsKey("excelStream")
                    || !o.fileName.ToString().EndsWith(".xlsx"))
                {
                    response.StatusCode = HttpStatusCode.BadRequest;
                    response.Headers.Add("UploadErrors", "{ \"Faltan datos\": \"Debe enviar mes(mm), gestion(yyyy), segmentoOrigen(id) y un archivo excel llamado file (en formato .xlsx)\"}");
                    response.Content = new StringContent("Debe enviar mes(mm), gestion(yyyy), segmentoOrigen(id) y un archivo excel llamado file");
                    return response;
                }

                int userid = Int32.Parse(Request.Headers.GetValues("id").First());
                var file = AddFileToProcess(o.mes.ToString(), o.gestion.ToString(), o.segmentoOrigen, ExcelFileType.Discount, userid, o.fileName.ToString());

                if (file == null)
                {
                    response.StatusCode = HttpStatusCode.BadRequest;
                    response.Headers.Add("UploadErrors", "{ \"Ya se Subio archivos para este mes\": \"Ya se subio  datos para este mes, si quiere volver a subir cancele el anterior archivo.\"}");
                    response.Content = new StringContent("Ya se subió  datos para este mes, si quiere volver a subir cancele el anterior archivo.");
                    return response;
                }

                ExcelFile = new DiscountExcel(o.excelStream, _context, o.fileName, o.mes, o.gestion,
                    o.segmentoOrigen.ToString(), file, headerin: ExcelHeaders, sheets: 1);
                if (ExcelFile.ValidateFile())
                {
                    ExcelFile.toDataBase();
                    file.State = FileState.UPLOADED;
                    _context.SaveChanges();
                    response.StatusCode = HttpStatusCode.OK;
                    response.Content = new StringContent("Se subio el archivo correctamente.");
                    return response;
                }

                file.State = FileState.CANCELED;
                _context.SaveChanges();
                return ExcelFile.toResponse();
            }
            catch (System.ArgumentException e)
            {
                response.StatusCode = HttpStatusCode.BadRequest;
                response.Headers.Add("UploadErrors", "{ \"Formato Archivo Invalido\": \"Por favor enviar un archivo en formato excel (.xlsx)\"}");
                ExcelFile.addError("Formato Archivo Invalido", "Por favor enviar un archivo en formato excel (.xlsx)");
                response.Content = new StringContent("Por favor enviar un archivo en formato excel (.xlsx)" + e);
                return response;
            }
            catch (System.IO.IOException e)
            {
                response.StatusCode = HttpStatusCode.BadRequest;
                response.Headers.Add("UploadErrors", "{ \"Archivo demasiado grande\": \"El archivo es demasiado grande para ser procesado.\"}");
                ExcelFile.addError("Archivo demasiado grande", "El archivo es demasiado grande para ser procesado.");
                response.Content = new StringContent("El archivo es demasiado grande para ser procesado.");
                return response;
            }
            catch (System.Exception e)
            {
                response.StatusCode = HttpStatusCode.BadRequest;
                response.Headers.Add("UploadErrors", "{ \"Existen Enlaces a otros archivos\": \"Existen celdas con referencias a otros archivos.\"}");
                ExcelFile.addError("Existen Enlaces a otros archivos", "Existen celdas con referencias a otros archivos.");
                response.Content = new StringContent("Por favor enviar un archivo en formato excel sin referencias a otros libros excel o formulas(.xls, .xslx)");
                return response;
            }
        }

        [HttpGet]
        [Route("api/payroll/PostgradoExcel")]
        public HttpResponseMessage GetPostgradoExcel()
        {
            PostgradoExcel contractExcel = new PostgradoExcel(fileName: "PosGrado.xlsx", headerin: ExcelHeaders);
            return contractExcel.getTemplate();
        }

        [HttpDelete]
        [Route("api/payroll/PostgradoExcel")]
        public IHttpActionResult CancelPostgradoExcel(JObject data)
        {
            int branchesid;
            if (data["mes"] == null || data["gestion"] == null || data["segmentoOrigen"] == null || !Int32.TryParse(data["segmentoOrigen"].ToString(), out branchesid))
            {
                ModelState.AddModelError("Mal Formato", "Debes enviar mes, gestion y segmentoOrigen");
                return BadRequest();

            }
            string mes = data["mes"].ToString();
            string gestion = data["gestion"].ToString();

            var file = _context.FileDbs.Include(f => f.DistProcess)
                .FirstOrDefault(f => f.DistProcess.mes == mes
                                     && f.DistProcess.gestion == gestion
                                     && f.DistProcess.BranchesId == branchesid
                                     && f.DistFileTypeId == (int) ExcelFileType.Postgrado
                                     && f.State == FileState.UPLOADED);
            if (file == null)
            {
                return NotFound();
            }

            file.State = FileState.CANCELED;
            _context.SaveChanges();

            return Ok();
        }

        [HttpPost]
        [Route("api/payroll/PostgradoExcel")]
        public async Task<HttpResponseMessage> UploadPostgradoExcel()
        {
            var response = new HttpResponseMessage();
            PostgradoExcel ExcelFile = null;
            try
            {
                var req = await Request.Content.ReadAsMultipartAsync();
                dynamic o = HttpContentToVariables(req).Result;

                if (!((IDictionary<string, object>)o).ContainsKey("mes")
                    || !((IDictionary<string, object>)o).ContainsKey("gestion")
                    || !((IDictionary<string, object>)o).ContainsKey("segmentoOrigen")
                    || !((IDictionary<string, object>)o).ContainsKey("fileName")
                    || !((IDictionary<string, object>)o).ContainsKey("excelStream")
                    || !o.fileName.ToString().EndsWith(".xlsx"))
                {
                    response.StatusCode = HttpStatusCode.BadRequest;
                    response.Headers.Add("UploadErrors", "{ \"Faltan datos\": \"Debe enviar mes(mm), gestion(yyyy), segmentoOrigen(id) y un archivo excel llamado file (en formato .xlsx)\"}");
                    response.Content = new StringContent("Debe enviar mes(mm), gestion(yyyy), segmentoOrigen(id) y un archivo excel llamado file");
                    return response;
                }

                int userid = Int32.Parse(Request.Headers.GetValues("id").First());
                var file = AddFileToProcess(o.mes.ToString(), o.gestion.ToString(), o.segmentoOrigen, ExcelFileType.Postgrado, userid, o.fileName.ToString());

                if (file == null)
                {
                    response.StatusCode = HttpStatusCode.BadRequest;
                    response.Headers.Add("UploadErrors", "{ \"Ya se Subio archivos para este mes\": \"Ya se subio  datos para este mes, si quiere volver a subir cancele el anterior archivo.\"}");
                    response.Content = new StringContent("Ya se subió  datos para este mes, si quiere volver a subir cancele el anterior archivo.");
                    return response;
                }

                ExcelFile = new PostgradoExcel(o.excelStream, _context, o.fileName, o.mes, o.gestion, o.segmentoOrigen.ToString(), file, headerin: ExcelHeaders, sheets: 1);
                if (ExcelFile.ValidateFile())
                {
                    ExcelFile.toDataBase();
                    file.State = FileState.UPLOADED;
                    _context.SaveChanges();
                    response.StatusCode = HttpStatusCode.OK;
                    response.Content = new StringContent("Se subio el archivo correctamente.");
                    return response;
                }
                file.State = FileState.ERROR;
                _context.SaveChanges();
                return ExcelFile.toResponse();
            }
            catch (System.ArgumentException e)
            {
                response.StatusCode = HttpStatusCode.BadRequest;
                response.Headers.Add("UploadErrors", "{ \"Formato Archivo Invalido\": \"Por favor enviar un archivo en formato excel (.xlsx)\"}");
                ExcelFile.addError("Formato Archivo Invalido", "Por favor enviar un archivo en formato excel (.xlsx)");
                response.Content = new StringContent("Por favor enviar un archivo en formato excel (.xlsx)" + e);
                return response;
            }
            catch (System.IO.IOException e)
            {
                response.StatusCode = HttpStatusCode.BadRequest;
                response.Headers.Add("UploadErrors", "{ \"Archivo demasiado grande\": \"El archivo es demasiado grande para ser procesado.\"}");
                ExcelFile.addError("Archivo demasiado grande", "El archivo es demasiado grande para ser procesado.");
                response.Content = new StringContent("El archivo es demasiado grande para ser procesado.");
                return response;
            }
            catch (System.Exception e)
            {
                response.StatusCode = HttpStatusCode.BadRequest;
                response.Headers.Add("UploadErrors", "{ \"Existen Enlaces a otros archivos\": \"Existen celdas con referencias a otros archivos.\"}");
                response.Content = new StringContent("Por favor enviar un archivo en formato excel sin referencias a otros libros excel o formulas(.xls, .xslx)");
                return response;
            }
        }

        [HttpGet]
        [Route("api/payroll/PregradoExcel")]
        public HttpResponseMessage GetPregradoExcel()
        {
            PregradoExcel contractExcel = new PregradoExcel(fileName: "Pregrado.xlsx", headerin: ExcelHeaders);
            return contractExcel.getTemplate();
        }

        [HttpDelete]
        [Route("api/payroll/PregradoExcel")]
        public IHttpActionResult CancelGetPregradoExcel(JObject data)
        {
            int branchesid;
            if (data["mes"] == null || data["gestion"] == null || data["segmentoOrigen"] == null || !Int32.TryParse(data["segmentoOrigen"].ToString(), out branchesid))
            {
                ModelState.AddModelError("Mal Formato", "Debes enviar mes, gestion y segmentoOrigen");
                return BadRequest();

            }
            string mes = data["mes"].ToString();
            string gestion = data["gestion"].ToString();
            var file = _context.FileDbs.Include(f => f.DistProcess)
                .FirstOrDefault(f => f.DistProcess.mes == mes
                                     && f.DistProcess.gestion == gestion
                                     && f.DistProcess.BranchesId == branchesid
                                     && f.DistFileTypeId == (int) ExcelFileType.Pregrado
                                     && f.State == FileState.UPLOADED);
            if (file == null)
            {
                return NotFound();
            }

            file.State = FileState.CANCELED;
            _context.SaveChanges();

            return Ok();
        }

        [HttpPost]
        [Route("api/payroll/PregradoExcel")]
        public async Task<HttpResponseMessage> UploadPregradoExcel()
        {
            var response = new HttpResponseMessage();
            PregradoExcel ExcelFile = null;
            try
            {
                var req = await Request.Content.ReadAsMultipartAsync();
                dynamic o = HttpContentToVariables(req).Result;

                if (!((IDictionary<string, object>)o).ContainsKey("mes")
                    || !((IDictionary<string, object>)o).ContainsKey("gestion")
                    || !((IDictionary<string, object>)o).ContainsKey("segmentoOrigen")
                    || !((IDictionary<string, object>)o).ContainsKey("fileName")
                    || !((IDictionary<string, object>)o).ContainsKey("excelStream")
                    || !o.fileName.ToString().EndsWith(".xlsx"))
                {
                    response.StatusCode = HttpStatusCode.BadRequest;
                    response.Headers.Add("UploadErrors", "{ \"Faltan datos\": \"Debe enviar mes(mm), gestion(yyyy), segmentoOrigen(id) y un archivo excel llamado file (en formato .xlsx)\"}");
                    response.Content = new StringContent("Debe enviar mes(mm), gestion(yyyy), segmentoOrigen(id) y un archivo excel llamado file");
                    return response;
                }

                int userid = Int32.Parse(Request.Headers.GetValues("id").First());
                var file = AddFileToProcess(o.mes.ToString(), o.gestion.ToString(), o.segmentoOrigen, ExcelFileType.Pregrado, userid, o.fileName.ToString());

                if (file == null)
                {
                    response.StatusCode = HttpStatusCode.BadRequest;
                    response.Headers.Add("UploadErrors", "{ \"Ya se Subio archivos para este mes\": \"Ya se subio  datos para este mes, si quiere volver a subir cancele el anterior archivo.\"}");
                    response.Content = new StringContent("Ya se subió  datos para este mes, si quiere volver a subir cancele el anterior archivo.");
                    return response;
                }

                ExcelFile = new PregradoExcel(o.excelStream, _context, o.fileName, o.mes, o.gestion, o.segmentoOrigen.ToString(), file, headerin: ExcelHeaders, sheets: 1);
                if (ExcelFile.ValidateFile())
                {
                    ExcelFile.toDataBase();
                    file.State = FileState.UPLOADED;
                    _context.SaveChanges();
                    response.StatusCode = HttpStatusCode.OK;
                    response.Content = new StringContent("Se subio el archivo correctamente.");
                    return response;
                }
                file.State = FileState.ERROR;
                _context.SaveChanges();
                return ExcelFile.toResponse();
            }
            catch (System.ArgumentException e)
            {
                response.StatusCode = HttpStatusCode.BadRequest;
                response.Headers.Add("UploadErrors", "{ \"Formato Archivo Invalido\": \"Por favor enviar un archivo en formato excel (.xlsx)\"}");
                ExcelFile.addError("Formato Archivo Invalido", "Por favor enviar un archivo en formato excel (.xlsx)");
                response.Content = new StringContent("Por favor enviar un archivo en formato excel (.xlsx)" + e);
                return response;
            }
            catch (System.IO.IOException e)
            {
                response.StatusCode = HttpStatusCode.BadRequest;
                response.Headers.Add("UploadErrors", "{ \"Archivo demasiado grande\": \"El archivo es demasiado grande para ser procesado.\"}");
                ExcelFile.addError("Archivo demasiado grande", "El archivo es demasiado grande para ser procesado.");
                response.Content = new StringContent("El archivo es demasiado grande para ser procesado.");
                return response;
            }
            catch (System.Exception e)
            {
                response.StatusCode = HttpStatusCode.BadRequest;
                response.Headers.Add("UploadErrors", "{ \"Existen Enlaces a otros archivos\": \"Existen celdas con referencias a otros archivos.\"}");
                response.Content = new StringContent("Por favor enviar un archivo en formato excel sin referencias a otros libros excel o formulas(.xls, .xslx)");
                return response;
            }
        }

        [HttpGet]
        [Route("api/payroll/ORExcel")]
        public HttpResponseMessage GetORExcel()
        {
            ORExcel contractExcel = new ORExcel(fileName: "OtrasRegionales.xlsx", headerin: ExcelHeaders);
            return contractExcel.getTemplate();
        }

        [HttpDelete]
        [Route("api/payroll/ORExcel")]
        public IHttpActionResult CancelORExcel(JObject data)
        {
            int branchesid;
            if (data["mes"] == null || data["gestion"] == null || data["segmentoOrigen"] == null || !Int32.TryParse(data["segmentoOrigen"].ToString(), out branchesid))
            {
                ModelState.AddModelError("Mal Formato", "Debes enviar mes, gestion y segmentoOrigen");
                return BadRequest();

            }
            string mes = data["mes"].ToString();
            string gestion = data["gestion"].ToString();
            var file = _context.FileDbs.Include(f => f.DistProcess)
                .FirstOrDefault(f => f.DistProcess.mes == mes
                                     && f.DistProcess.gestion == gestion
                                     && f.DistProcess.BranchesId == branchesid
                                     && f.DistFileTypeId == (int) ExcelFileType.OR
                                     && f.State == FileState.UPLOADED);
            if (file == null)
            {
                return NotFound();
            }

            file.State = FileState.CANCELED;
            _context.SaveChanges();

            return Ok();
        }

        [HttpPost]
        [Route("api/payroll/ORExcel")]
        public async Task<HttpResponseMessage> UploadORExcel()
        {
            var response = new HttpResponseMessage();
            ORExcel ExcelFile = null;
            try
            {
                var req = await Request.Content.ReadAsMultipartAsync();
                dynamic o = HttpContentToVariables(req).Result;

                if (!((IDictionary<string, object>)o).ContainsKey("mes")
                    || !((IDictionary<string, object>)o).ContainsKey("gestion")
                    || !((IDictionary<string, object>)o).ContainsKey("segmentoOrigen")
                    || !((IDictionary<string, object>)o).ContainsKey("fileName")
                    || !((IDictionary<string, object>)o).ContainsKey("excelStream")
                    || !o.fileName.ToString().EndsWith(".xlsx"))
                {
                    response.StatusCode = HttpStatusCode.BadRequest;
                    response.Headers.Add("UploadErrors", "{ \"Faltan datos\": \"Debe enviar mes(mm), gestion(yyyy), segmentoOrigen(id) y un archivo excel llamado file (en formato .xlsx)\"}");
                    response.Content = new StringContent("Debe enviar mes(mm), gestion(yyyy), segmentoOrigen(id) y un archivo excel llamado file");
                    return response;
                }

                int userid = Int32.Parse(Request.Headers.GetValues("id").First());
                var file = AddFileToProcess(o.mes.ToString(), o.gestion.ToString(), o.segmentoOrigen, ExcelFileType.OR, userid, o.fileName.ToString());

                if (file == null)
                {
                    response.StatusCode = HttpStatusCode.BadRequest;
                    response.Headers.Add("UploadErrors", "{ \"Ya se Subio archivos para este mes\": \"Ya se subio  datos para este mes, si quiere volver a subir cancele el anterior archivo.\"}");
                    response.Content = new StringContent("Ya se subió  datos para este mes, si quiere volver a subir cancele el anterior archivo.");
                    return response;
                }

                ExcelFile = new ORExcel(o.excelStream, _context, o.fileName, o.mes, o.gestion, o.segmentoOrigen.ToString(), file, headerin: ExcelHeaders, sheets: 1);
                if (ExcelFile.ValidateFile())
                {
                    ExcelFile.toDataBase();
                    file.State = FileState.UPLOADED;
                    _context.SaveChanges();
                    response.StatusCode = HttpStatusCode.OK;
                    response.Content = new StringContent("Se subio el archivo correctamente.");
                    return response;
                }
                file.State = FileState.ERROR;
                _context.SaveChanges();
                return ExcelFile.toResponse();
            }
            catch (System.ArgumentException e)
            {
                response.StatusCode = HttpStatusCode.BadRequest;
                response.Headers.Add("UploadErrors", "{ \"Formato Archivo Invalido\": \"Por favor enviar un archivo en formato excel (.xlsx)\"}");
                ExcelFile.addError("Formato Archivo Invalido", "Por favor enviar un archivo en formato excel (.xlsx)");
                response.Content = new StringContent("Por favor enviar un archivo en formato excel (.xlsx)" + e);
                return response;
            }
            catch (System.IO.IOException e)
            {
                response.StatusCode = HttpStatusCode.BadRequest;
                response.Headers.Add("UploadErrors", "{ \"Archivo demasiado grande\": \"El archivo es demasiado grande para ser procesado.\"}");
                ExcelFile.addError("Archivo demasiado grande", "El archivo es demasiado grande para ser procesado.");
                response.Content = new StringContent("El archivo es demasiado grande para ser procesado.");
                return response;
            }
            catch (System.Exception e)
            {
                response.StatusCode = HttpStatusCode.BadRequest;
                response.Headers.Add("UploadErrors", "{ \"Existen Enlaces a otros archivos\": \"Existen celdas con referencias a otros archivos.\"}");
                response.Content = new StringContent("Por favor enviar un archivo en formato excel sin referencias a otros libros excel o formulas(.xls, .xslx)");
                return response;
            }
        }

        [HttpGet]
        [Route("api/payroll/GetErrors/{id}")]
        public IHttpActionResult GetErrors(int id)
        {
            var process = _context.DistProcesses.FirstOrDefault(p => p.Id == id);
            
            if (process == null)
                return NotFound();

            if (process.State != ProcessState.ERROR && process.State != ProcessState.WARNING)
                return Ok();

            var err = _context.DistLogErroreses.Where(e => e.DistProcessId == process.Id && !e.Inspected ).Include(e=>e.Error).Select(e=>new{e.Id,e.ErrorId,e.Error.Name,e.Error.Description,e.Error.Type,e.Archivos,e.CUNI});
            return Ok(err);
        }


        [HttpGet]
        [Route("api/payroll/Validate/{id}")]
        public IHttpActionResult Validate(int id)
        {
            var process = _context.DistProcesses.FirstOrDefault(p => p.Id == id && p.State == ProcessState.STARTED);
            if (process == null)
                return NotFound();

            int userid = Int32.Parse(Request.Headers.GetValues("id").First());
            _context.Database.ExecuteSqlCommand("CALL ADMNALRRHH.FIX_ACAD(" + process.Id + ")");

            _context.Database.ExecuteSqlCommand("CALL ADMNALRRHH.VALIDATE_HASALLFILES(" + userid + "," + process.Id + ")");
            _context.Database.ExecuteSqlCommand("CALL ADMNALRRHH.VALIDATE_TIPOEMPLEADO(" + userid + "," + process.Id + ")");
            _context.Database.ExecuteSqlCommand("CALL ADMNALRRHH.VALIDATE_CE(" + userid + "," + process.Id + ")");
            _context.Database.ExecuteSqlCommand("CALL ADMNALRRHH.VALIDATE_OD(" + userid + "," + process.Id + ")");
            _context.Database.ExecuteSqlCommand("CALL ADMNALRRHH.VALIDATE_OR(" + userid + "," + process.Id + ")");

            _context.Database.ExecuteSqlCommand("CALL ADMNALRRHH.VALIDATE_CUADRARDESCUENTOS(" + userid + "," + process.Id + ")");
            _context.Database.ExecuteSqlCommand("CALL ADMNALRRHH.VALIDATE_ACADSUM(" + userid + "," + process.Id + ")");
            _context.Database.ExecuteSqlCommand("CALL ADMNALRRHH.VALIDATE_OTHERINCOMES(" + userid + "," + process.Id + ")");


            var err = _context.DistLogErroreses.Where(e => e.DistProcessId == process.Id && !e.Inspected).Include(e => e.Error).Select(e => new { e.Id, e.ErrorId, e.Error.Name, e.Error.Description,e.Error.Type, e.Archivos, e.CUNI });
            if (err.Count()>0)
            {
                if(err.Where(e => e.Type == "E").Count()>0)
                    process.State = ProcessState.ERROR;
                else
                    process.State = ProcessState.WARNING;
                _context.SaveChanges();
                return Ok("Se encontró errores en los archivos subidos");
            }

            process.State = ProcessState.VALIDATED;
            _context.SaveChanges();
            return Ok("La información es correcta");
            
        }

        [HttpGet]
        [Route("api/payroll/AcceptWarnings/{id}")]
        public IHttpActionResult AcceptWarnings(int id)
        {
            var process = _context.DistProcesses.FirstOrDefault(p => p.Id == id && p.State==ProcessState.WARNING);
            if (process == null)
                return NotFound();
            process.State = ProcessState.VALIDATED;
            _context.SaveChanges();
            return Ok();
        }

        [HttpGet]
        [Route("api/payroll/Distribute/{id}")]
        public IHttpActionResult Distribute(int id)
        {
            var process = _context.DistProcesses.FirstOrDefault(p => p.Id == id && p.State == ProcessState.VALIDATED);
            if (process == null)
                return NotFound();

            _context.Database.ExecuteSqlCommand("CALL ADMNALRRHH.SET_PERCENTS(" + process.Id + ")");
            _context.Database.ExecuteSqlCommand("CALL ADMNALRRHH.DIST_PERCENTS(" + process.Id + ")");
            _context.Database.ExecuteSqlCommand("CALL ADMNALRRHH.DIST_COSTS(" + process.Id + ")");
            process.State = ProcessState.PROCESSED;
            _context.SaveChanges();

            return Ok("Se procesó la información");
        }

        [HttpGet]
        [Route("api/payroll/Process")]
        public IHttpActionResult GetProcesses()
        {
            var processes = _context.DistProcesses.Include(p => p.Branches)
                .Where(p => p.State != ProcessState.CANCELED).
                Select(p=> new
                {
                    p.BranchesId,
                    p.Branches.Name,
                    p.State,
                    p.Id,
                    p.gestion,
                    p.mes
                });

            var user = auth.getUser(Request);
            var res = auth.filerByRegional(processes, user);

            return Ok(res);
        }

        [HttpDelete]
        [Route("api/payroll/Process/{id}")]
        public IHttpActionResult Process(int id)
        {

            var processInDB = _context.DistProcesses.FirstOrDefault(p => p.Id == id && (p.State != ProcessState.CANCELED || p.State != ProcessState.INSAP ));
            if (processInDB == null)
                return NotFound();
            processInDB.State = ProcessState.CANCELED;
            _context.SaveChanges();

            var files = _context.FileDbs.Where(f => f.State == FileState.UPLOADED && f.DistProcessId==processInDB.Id);
            foreach (var file in files)
            {
                file.State = FileState.CANCELED;
            }
            _context.SaveChanges();
            return Ok("Proceso Cancelado");
        }

        [HttpGet]
        [Route("api/payroll/GetDistribution/{id}")]
        public HttpResponseMessage GetDistribution(int id)
        {
            HttpResponseMessage response = new HttpResponseMessage();

            var pro = _context.DistProcesses.Include(x => x.Branches).FirstOrDefault(x => x.Id == id);
            if (pro == null)
            {
                response.StatusCode = HttpStatusCode.NotFound;
                return response;
            }

            IEnumerable<Distribution> dist = _context.Database.SqlQuery<Distribution>("SELECT a.\"Document\",a.\"TipoEmpleado\",a.\"Dependency\",a.\"PEI\","+
            " a.\"PlanEstudios\",a.\"Paralelo\",a.\"Periodo\",a.\"Project\",a.\"BussinesPartner\","+
            " a.\"Monto\",a.\"Porcentaje\",a.\"MontoDividido\",a.\"segmentoOrigen\",a.\"BussinesPartner\"," +
            " b.\"mes\",b.\"gestion\",e.\"Name\" as Segmento ,d.\"Concept\",d.\"Name\" as CuentasContables,d.\"Indicator\" " +
            " FROM ADMNALRRHH.\"Dist_Cost\" a "+
                " INNER JOIN  ADMNALRRHH.\"Dist_Process\" b " + 
                " on a.\"DistProcessId\"=b.\"Id\" "+
            " AND a.\"DistProcessId\"= "+ id +
            " INNER JOIN  ADMNALRRHH.\"Dist_TipoEmpleado\" c "+
                "on a.\"TipoEmpleado\"=c.\"Name\" "+
            " INNER JOIN  ADMNALRRHH.\"CuentasContables\" d "+
               " on c.\"GrupoContableId\" = d.\"GrupoContableId\" " +
            " and b.\"BranchesId\" = d.\"BranchesId\" "+
            " and a.\"Columna\" = d.\"Concept\" "+
            " INNER JOIN ADMNALRRHH.\"Branches\" e "+
               " on b.\"BranchesId\" = e.\"Id\"").ToList();

            var ex = new XLWorkbook();
            var d = new Distribution();
            ex.Worksheets.Add(d.CreateDataTable(dist), "Detalle");


            var ms = new MemoryStream();
            ex.SaveAs(ms);
            response.StatusCode = HttpStatusCode.OK;
            response.Content = new StreamContent(ms);
            response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");
            response.Content.Headers.ContentDisposition.FileName = "Distribucion-" + pro.Branches.Abr + "-" + pro.mes + pro.gestion + ".xlsx";
            response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            response.Content.Headers.ContentLength = ms.Length;
            ms.Seek(0, SeekOrigin.Begin); 
            return response;
        }

        [HttpGet]
        [Route("api/payroll/GetTotalGeneral/{id}")]
        public HttpResponseMessage GetTotalGeneral(int id)
        {
            HttpResponseMessage response = new HttpResponseMessage();

            var pro = _context.DistProcesses.Include(x => x.Branches).FirstOrDefault(x => x.Id == id);
            if (pro == null)
            {
                response.StatusCode = HttpStatusCode.NotFound;
                return response;
            }

            IEnumerable<Distribution> dist = _context.Database.SqlQuery<Distribution>("SELECT a.\"Document\",a.\"TipoEmpleado\",a.\"Dependency\",a.\"PEI\"," +
            " a.\"PlanEstudios\",a.\"Paralelo\",a.\"Periodo\",a.\"Project\",a.\"BussinesPartner\"," +
            " a.\"Monto\",a.\"Porcentaje\",a.\"MontoDividido\",a.\"segmentoOrigen\",a.\"BussinesPartner\"," +
            " b.\"mes\",b.\"gestion\",e.\"Name\" as Segmento ,d.\"Concept\",d.\"Name\" as CuentasContables,d.\"Indicator\" " +
            " FROM ADMNALRRHH.\"Dist_Cost\" a " +
                " INNER JOIN  ADMNALRRHH.\"Dist_Process\" b " +
                " on a.\"DistProcessId\"=b.\"Id\" " +
            " AND a.\"DistProcessId\"= " + id +
            " INNER JOIN  ADMNALRRHH.\"Dist_TipoEmpleado\" c " +
                "on a.\"TipoEmpleado\"=c.\"Name\" " +
            " INNER JOIN  ADMNALRRHH.\"CuentasContables\" d " +
               " on c.\"GrupoContableId\" = d.\"GrupoContableId\" " +
            " and b.\"BranchesId\" = d.\"BranchesId\" " +
            " and a.\"Columna\" = d.\"Concept\" " +
            " INNER JOIN ADMNALRRHH.\"Branches\" e " +
               " on b.\"BranchesId\" = e.\"Id\"").ToList();

            var groupedD = dist.Where(x=>x.Indicator=="D").GroupBy(x => new
                                            {
                                                x.Concept, x.Indicator
                                            })
                                .Select(y=> new
                                            {
                                                Concepto = y.First().Concept,
                                                D = y.Sum(z => z.MontoDividido),
                                                H = 0.0m
                                            });
            var groupedH = dist.Where(x => x.Indicator == "H").GroupBy(x => new
                {
                    x.Concept,
                    x.Indicator
                })
                .Select(y => new
                {
                    Concepto = y.First().Concept,
                    D = 0.0m,
                    H = y.Sum(z => z.MontoDividido)
                });
            var res = groupedH.Concat(groupedD);

            var ex = new XLWorkbook();
            var d = new Distribution();
            ex.Worksheets.Add(d.CreateDataTable(res), "Detalle");


            var ms = new MemoryStream();
            ex.SaveAs(ms);
            response.StatusCode = HttpStatusCode.OK;
            response.Content = new StreamContent(ms);
            response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");
            response.Content.Headers.ContentDisposition.FileName = "TotalGeneral-" + pro.Branches.Abr + "-" + pro.mes + pro.gestion + ".xlsx";
            response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            response.Content.Headers.ContentLength = ms.Length;
            ms.Seek(0, SeekOrigin.Begin);
            return response;
        }

        [HttpGet]
        [Route("api/payroll/GetTotalCuenta/{id}")]
        public HttpResponseMessage GetTotalCuenta(int id)
        {
            HttpResponseMessage response = new HttpResponseMessage();

            var pro = _context.DistProcesses.Include(x => x.Branches).FirstOrDefault(x => x.Id == id);
            if (pro == null)
            {
                response.StatusCode = HttpStatusCode.NotFound;
                return response;
            }

            IEnumerable<Distribution> dist = _context.Database.SqlQuery<Distribution>("SELECT a.\"Document\",a.\"TipoEmpleado\",a.\"Dependency\",a.\"PEI\"," +
            " a.\"PlanEstudios\",a.\"Paralelo\",a.\"Periodo\",a.\"Project\",a.\"BussinesPartner\"," +
            " a.\"Monto\",a.\"Porcentaje\",a.\"MontoDividido\",a.\"segmentoOrigen\",a.\"BussinesPartner\"," +
            " b.\"mes\",b.\"gestion\",e.\"Name\" as Segmento ,d.\"Concept\",d.\"Name\" as CuentasContables,d.\"Indicator\" " +
            " FROM ADMNALRRHH.\"Dist_Cost\" a " +
                " INNER JOIN  ADMNALRRHH.\"Dist_Process\" b " +
                " on a.\"DistProcessId\"=b.\"Id\" " +
            " AND a.\"DistProcessId\"= " + id +
            " INNER JOIN  ADMNALRRHH.\"Dist_TipoEmpleado\" c " +
                "on a.\"TipoEmpleado\"=c.\"Name\" " +
            " INNER JOIN  ADMNALRRHH.\"CuentasContables\" d " +
               " on c.\"GrupoContableId\" = d.\"GrupoContableId\" " +
            " and b.\"BranchesId\" = d.\"BranchesId\" " +
            " and a.\"Columna\" = d.\"Concept\" " +
            " INNER JOIN ADMNALRRHH.\"Branches\" e " +
               " on b.\"BranchesId\" = e.\"Id\"").ToList();

            var groupedD = dist.Where(x => x.Indicator == "D").GroupBy(x => new
                {
                    x.CuentasContables,
                    x.Indicator
                })
                .Select(y => new
                {
                    Cuenta = y.First().CuentasContables,
                    D = y.Sum(z => z.MontoDividido),
                    H = 0.0m
                });
            var groupedH = dist.Where(x => x.Indicator == "H").GroupBy(x => new
                {
                    x.CuentasContables,
                    x.Indicator
                })
                .Select(y => new
                {
                    Cuenta = y.First().CuentasContables,
                    D = 0.0m,
                    H = y.Sum(z => z.MontoDividido)
                });
            var res = groupedH.Concat(groupedD);

            var ex = new XLWorkbook();
            var d = new Distribution();
            ex.Worksheets.Add(d.CreateDataTable(res), "Detalle");


            var ms = new MemoryStream();
            ex.SaveAs(ms);
            response.StatusCode = HttpStatusCode.OK;
            response.Content = new StreamContent(ms);
            response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");
            response.Content.Headers.ContentDisposition.FileName = "TotalPorCuenta-" + pro.Branches.Abr + "-" + pro.mes + pro.gestion + ".xlsx";
            response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            response.Content.Headers.ContentLength = ms.Length;
            ms.Seek(0, SeekOrigin.Begin);
            return response;
        }


        [HttpGet]
        [Route("api/payroll/GetSAPResume/{id}")]
        public HttpResponseMessage GetSAPResume(int id)
        {
            HttpResponseMessage response = new HttpResponseMessage();
            
            var pro = _context.DistProcesses.Include(x=>x.Branches).FirstOrDefault(x => x.Id == id);
            if (pro == null)
            {
                response.StatusCode = HttpStatusCode.NotFound;
                return response;
            }

            IEnumerable<SapVoucher> dist = _context.Database.SqlQuery<SapVoucher>("SELECT null \"ParentKey\",null \"LineNum\",null \"AccountCode\",null \"Debit\",null \"Credit\",null \"ShortName\", null as \"LineMemo\",null \"ProjectCode\",null \"CostingCode\",null \"CostingCode2\",null \"CostingCode3\",null \"CostingCode4\",null \"CostingCode5\" from dummy " +
                                                                                  " union  SELECT \"ParentKey\",\"LineNum\",\"AccountCode\",case when replace(sum(\"Debit\"),',','.')='0.00' then null else replace(sum(\"Debit\"),',','.') end \"Debit\",case when replace(sum(\"Credit\"),',','.')='0.00' then null else replace(sum(\"Credit\"),',','.') end \"Credit\", \"ShortName\", null as \"LineMemo\",\"ProjectCode\",\"CostingCode\",\"CostingCode2\",\"CostingCode3\",\"CostingCode4\",\"CostingCode5\" " +
                                                                                        " FROM ("+
                                                                                        " select 1 \"ParentKey\"," +
                                                                                        "  null \"LineNum\"," +
                                                                                        "  coalesce(b.\"AcctCode\",x.\"CUENTASCONTABLES\") \"AccountCode\"," +
                                                                                        "  CASE WHEN x.\"Indicator\"='D' then x.\"MontoDividido\" else 0 end as \"Debit\"," +
                                                                                        "  CASE WHEN x.\"Indicator\"='H' then x.\"MontoDividido\"else 0 end as \"Credit\"," +
                                                                                        "  x.\"BussinesPartner\" \"ShortName\"," +
                                                                                        "  x.\"Concept\" \"LineMemo\"," +
                                                                                        "  x.\"Project\" \"ProjectCode\"," +
                                                                                        "  f.\"Cod\" \"CostingCode\"," +
                                                                                        "  x.\"PEI\" \"CostingCode2\"," +
                                                                                        "  x.\"PlanEstudios\" \"CostingCode3\"," +
                                                                                        "  x.\"Paralelo\" \"CostingCode4\"," +
                                                                                        "  x.\"Periodo\" \"CostingCode5\"" +
                                                                                        " from  (SELECT a.\"Document\",a.\"TipoEmpleado\",a.\"Dependency\",a.\"PEI\","+
                                                                                        "           a.\"PlanEstudios\",a.\"Paralelo\",a.\"Periodo\",a.\"Project\","+
                                                                                        "           a.\"Monto\",a.\"Porcentaje\",a.\"MontoDividido\",a.\"segmentoOrigen\",a.\"BussinesPartner\","+
                                                                                        "           b.\"mes\",b.\"gestion\",e.\"Name\" as Segmento ,d.\"Concept\",d.\"Name\" as CuentasContables,d.\"Indicator\""+
                                                                                        "           FROM ADMNALRRHH.\"Dist_Cost\" a "+
                                                                                        "               INNER JOIN  ADMNALRRHH.\"Dist_Process\" b "+
                                                                                        "               on a.\"DistProcessId\"=b.\"Id\" "+
                                                                                        "           AND a.\"DistProcessId\"= " + id +
                                                                                        "           INNER JOIN  ADMNALRRHH.\"Dist_TipoEmpleado\" c "+
                                                                                        "                on a.\"TipoEmpleado\"=c.\"Name\" "+
                                                                                        "           INNER JOIN  ADMNALRRHH.\"CuentasContables\" d "+
                                                                                        "              on c.\"GrupoContableId\" = d.\"GrupoContableId\""+
                                                                                        "           and b.\"BranchesId\" = d.\"BranchesId\" "+
                                                                                        "           and a.\"Columna\" = d.\"Concept\" "+
                                                                                        "           INNER JOIN ADMNALRRHH.\"Branches\" e "+
                                                                                        "              on b.\"BranchesId\" = e.\"Id\") x"+
                                                                                        " left join ucatolica.oact b"+
                                                                                        " on x.CUENTASCONTABLES=b.\"FormatCode\""+
                                                                                        " left join admnalrrhh.\"Dependency\" d"+
                                                                                        " on x.\"Dependency\"=d.\"Cod\""+
                                                                                        " left join admnalrrhh.\"OrganizationalUnit\" f"+
                                                                                        " on d.\"OrganizationalUnitId\"=f.\"Id\""+
                                                                                        ") V "+
                                                                                        "GROUP BY \"ParentKey\",\"LineNum\",\"AccountCode\", \"ShortName\",\"ProjectCode\",\"CostingCode\",\"CostingCode2\",\"CostingCode3\",\"CostingCode4\",\"CostingCode5\";").ToList();

            var ex = new XLWorkbook();
            var d = new Distribution();

            var lastday = pro.gestion + pro.mes + DateTime.DaysInMonth(Int32.Parse(pro.gestion), Int32.Parse(pro.mes)).ToString();

            IEnumerable<VoucherHeader> dist1 = _context.Database.SqlQuery<VoucherHeader>("SELECT null \"ParentKey\", null \"ReferenceDate\",null \"Memo\",null \"TaxDate\",null \"Series\",null \"DueDate\" FROM DUMMY " +
                                                                                         "union SELECT '1' \"ParentKey\", '" + lastday + "' \"ReferenceDate\",'Planilla Menusal " + pro.Branches.Abr + "-" + pro.mes + "-" + pro.gestion + "' \"Memo\",'" + lastday + "' \"TaxDate\",'" + pro.Branches.SerieComprobanteContalbeSAP + "' \"Series\",'" + lastday + "' \"DueDate\" FROM DUMMY;");
            var n = d.CreateDataTable(dist1);
            int desiredSize = 1;

            while (n.Columns.Count > desiredSize)
            {
                n.Columns.RemoveAt(desiredSize);
            }
            ex.Worksheets.Add(n,"Voucher");

            ex.Worksheets.Add(d.CreateDataTable(dist1), "Cabecera");
            
            ex.Worksheets.Add(d.CreateDataTable(dist), "Detalle");

            var ms = new MemoryStream();
            ex.SaveAs(ms);
            response.StatusCode = HttpStatusCode.OK;
            response.Content = new StreamContent(ms);
            response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");
            response.Content.Headers.ContentDisposition.FileName = "SAP_Voucher_Lines-"+pro.Branches.Abr+"-"+pro.mes+pro.gestion+".xlsx";
            response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            response.Content.Headers.ContentLength = ms.Length;
            ms.Seek(0, SeekOrigin.Begin); 
            return response;
        }

        [HttpGet]
        [Route("api/payroll/SumTotalesPlanilla/{id}")]
        public IHttpActionResult SumTotalesPlanilla(int id)
        {
            var process = _context.DistProcesses.FirstOrDefault(p => p.Id == id && p.State == ProcessState.VALIDATED);

            if (process == null)
                return NotFound();

            var file = _context.FileDbs.FirstOrDefault(f => f.DistFileTypeId == 1 && f.DistProcessId==process.Id  && f.State == FileState.UPLOADED);

            if (file == null)
                return NotFound();

            var sum = _context.DistPayrolls.Where(p => p.DistFileId == file.Id).Select(
                p=>new
                {
                    p.BasicSalary,
                    p.AntiquityBonus,
                    p.OtherIncome,
                    p.TeachingIncome,
                    p.OtherAcademicIncomes,
                    p.Reintegro,
                    p.TotalAmountEarned,
                    p.AFPLaboral,
                    p.RcIva,
                    p.Discounts,
                    p.TotalAmountDiscounts,
                    p.TotalAfterDiscounts,
                    p.AFPPatronal,
                    p.SeguridadCortoPlazoPatronal,
                    p.ProvAguinaldo,
                    p.ProvPrimas,
                    p.ProvIndeminizacion
                });

            List<string> names = new List<string>()
            {
                "Haber Basico",
                "Bono Antiguedad",
                "Otros Ingresos",
                "Ingresos por Docencia",
                "Ingresos por otras actividades academicas",
                "Reintegro",
                "Total Ganado",
                "Aporte Laboral AFP",
                "RC IVA",
                "Descuentos",
                "Total Deducciones",
                "Liquido Pagable",
                "Aporte patronal AFP",
                "Aporte patronal SCP",
                "Provision Aguinaldos",
                "Provision Primas",
                "Provision Indeminizacion",
            };
            List<JObject> result = new List<JObject>();

            List<decimal> totales = new List<decimal>()
            {
                sum.Sum(p => p.BasicSalary),
                sum.Sum(p => p.AntiquityBonus),
                sum.Sum(p => p.OtherIncome),
                sum.Sum(p => p.TeachingIncome),
                sum.Sum(p => p.OtherAcademicIncomes),
                sum.Sum(p => p.Reintegro),
                sum.Sum(p => p.TotalAmountEarned),
                sum.Sum(p => p.AFPLaboral),
                sum.Sum(p => p.RcIva),
                sum.Sum(p => p.Discounts),
                sum.Sum(p => p.TotalAmountDiscounts),
                sum.Sum(p => p.TotalAfterDiscounts),
                sum.Sum(p => p.AFPPatronal),
                sum.Sum(p => p.SeguridadCortoPlazoPatronal),
                sum.Sum(p => p.ProvAguinaldo),
                sum.Sum(p => p.ProvPrimas),
                sum.Sum(p => p.ProvIndeminizacion),
            };

            for (int j = 0; j < totales.Count; j++)
            {
                dynamic re = new JObject();
                re.name = names[j];
                re.total = totales[j];
                result.Add(re);
            }
            

            return Ok(result);
        }

    }



}
