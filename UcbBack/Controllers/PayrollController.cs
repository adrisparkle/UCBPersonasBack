using System;
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

        public PayrollController()
        {
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
            var processInDB =
                _context.DistProcesses.FirstOrDefault(p =>
                    p.BranchesId == BranchesId && p.gestion == gestion && p.mes == mes && (p.State == ProcessState.STARTED || p.State == ProcessState.ERROR));

            if (processInDB != null)
            {
                var fileInDB = _context.FileDbs.FirstOrDefault(f => f.DistProcessId == processInDB.Id && f.DistFileTypeId == (int)FileType && f.State == FileState.UPLOADED);
                if (fileInDB == null)
                {
                    processInDB.State = ProcessState.STARTED;
                    _context.Database.ExecuteSqlCommand("UPDATE ADMNALRRHH.\"Dist_LogErrores\" set \"Inspected\" = true where \"DistProcessId\" = " + processInDB.Id);
                    var file = new Dist_File();
                    file.Id = _context.Database.SqlQuery<int>("SELECT \"rrhh_Dist_File_sqs\".nextval FROM DUMMY;").ToList()[0];
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
                process.Id = _context.Database.SqlQuery<int>("SELECT \"rrhh_Dist_Process_sqs\".nextval FROM DUMMY;").ToList()[0];
                process.BranchesId = BranchesId;
                process.mes = mes;
                process.gestion = gestion;
                process.State = ProcessState.STARTED;
                _context.DistProcesses.Add(process);
                _context.SaveChanges();

                var file = new Dist_File();
                file.Id = _context.Database.SqlQuery<int>("SELECT \"rrhh_Dist_File_sqs\".nextval FROM DUMMY;").ToList()[0];
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
            try
            {
                var req = await Request.Content.ReadAsMultipartAsync();
                dynamic o = HttpContentToVariables(req).Result;

                if (!((IDictionary<string, object>)o).ContainsKey("mes") 
                    || !((IDictionary<string, object>)o).ContainsKey("gestion") 
                    || !((IDictionary<string, object>)o).ContainsKey("segmentoOrigen")
                    || !((IDictionary<string, object>)o).ContainsKey("fileName")
                    || !((IDictionary<string, object>)o).ContainsKey("excelStream"))
                {
                    response.StatusCode = HttpStatusCode.BadRequest;
                    response.Content = new StringContent("Debe enviar mes(mm), gestion(yyyy), segmentoOrigen(id) y un archivo excel llamado file");
                    return response;
                }

                int userid = Int32.Parse(Request.Headers.GetValues("id").First());
                var file = AddFileToProcess(o.mes.ToString(), o.gestion.ToString(), o.segmentoOrigen, ExcelFileType.Payroll, userid, o.fileName.ToString());

                if (file == null)
                {
                    response.StatusCode = HttpStatusCode.BadRequest;
                    response.Content = new StringContent("Ya se subió  datos para este mes, si quiere volver a subir cancele el anterior archivo.");
                    return response;
                }

                PayrollExcel contractExcel = new PayrollExcel(o.excelStream, _context, o.fileName, o.mes, o.gestion, o.segmentoOrigen.ToString(), file, headerin: ExcelHeaders, sheets: 1);
                if (contractExcel.ValidateFile())
                {
                    contractExcel.toDataBase();
                    file.State = FileState.UPLOADED;
                    _context.SaveChanges();
                    response.StatusCode = HttpStatusCode.OK;
                    response.Content = new StringContent("Se subio el archivo correctamente.");
                    return response;
                }
                file.State = FileState.ERROR;
                _context.SaveChanges();
                return contractExcel.toResponse();
            }
            catch (System.ArgumentException e)
            {
                response.StatusCode = HttpStatusCode.BadRequest;
                response.Content = new StringContent("Por favor enviar un archivo en formato excel (.xls, .xslx)" + e);
                return response;
            }
            catch (System.IO.IOException e)
            {
                response.StatusCode = HttpStatusCode.BadRequest;
                response.Content = new StringContent("El archivo es demasiado grande para ser procesado.");
                return response;
            }
            catch (System.Exception e)
            {
                response.StatusCode = HttpStatusCode.BadRequest;
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
            try
            {
                var req = await Request.Content.ReadAsMultipartAsync();
                dynamic o = HttpContentToVariables(req).Result;

                if (!((IDictionary<string, object>)o).ContainsKey("mes")
                    || !((IDictionary<string, object>)o).ContainsKey("gestion")
                    || !((IDictionary<string, object>)o).ContainsKey("segmentoOrigen")
                    || !((IDictionary<string, object>)o).ContainsKey("fileName")
                    || !((IDictionary<string, object>)o).ContainsKey("excelStream"))
                {
                    response.StatusCode = HttpStatusCode.BadRequest;
                    response.Content = new StringContent("Debe enviar mes(mm), gestion(yyyy), segmentoOrigen(id) y un archivo excel llamado file");
                    return response;
                }

                int userid = Int32.Parse(Request.Headers.GetValues("id").First());
                var file = AddFileToProcess(o.mes, o.gestion, o.segmentoOrigen, ExcelFileType.Academic, userid, o.fileName);

                if (file == null)
                {
                    response.StatusCode = HttpStatusCode.BadRequest;
                    response.Content = new StringContent("Ya se subió  datos para este mes, si quiere volver a subir cancele el anterior archivo.");
                    return response;
                }

                AcademicExcel contractExcel = new AcademicExcel(o.excelStream, _context, o.fileName, o.mes, o.gestion, o.segmentoOrigen.ToString(), file, headerin: ExcelHeaders, sheets: 1);
                if (contractExcel.ValidateFile())
                {
                    contractExcel.toDataBase();
                    response.StatusCode = HttpStatusCode.OK;
                    file.State = FileState.UPLOADED;
                    _context.SaveChanges();
                    response.Content = new StringContent("Se subio el archivo correctamente.");
                    return response;
                }
                file.State = FileState.ERROR;
                _context.SaveChanges();
                return contractExcel.toResponse();
            }
            catch (System.ArgumentException e)
            {
                response.StatusCode = HttpStatusCode.BadRequest;
                response.Content = new StringContent("Por favor enviar un archivo en formato excel (.xls, .xslx)" + e);
                return response;
            }
            catch (System.IO.IOException e)
            {
                response.StatusCode = HttpStatusCode.BadRequest;
                response.Content = new StringContent("El archivo es demasiado grande para ser procesado.");
                return response;
            }
            catch (System.Exception e)
            {
                response.StatusCode = HttpStatusCode.BadRequest;
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
            try
            {
                var req = await Request.Content.ReadAsMultipartAsync();
                dynamic o = HttpContentToVariables(req).Result;

                if (!((IDictionary<string, object>) o).ContainsKey("mes")
                    || !((IDictionary<string, object>) o).ContainsKey("gestion")
                    || !((IDictionary<string, object>) o).ContainsKey("segmentoOrigen")
                    || !((IDictionary<string, object>) o).ContainsKey("fileName")
                    || !((IDictionary<string, object>) o).ContainsKey("excelStream"))
                {
                    response.StatusCode = HttpStatusCode.BadRequest;
                    response.Content =
                        new StringContent(
                            "Debe enviar mes(mm), gestion(yyyy), segmentoOrigen(id) y un archivo excel llamado file");
                    return response;
                }

                int userid = Int32.Parse(Request.Headers.GetValues("id").First());
                var file = AddFileToProcess(o.mes, o.gestion, o.segmentoOrigen, ExcelFileType.Discount, userid,
                    o.fileName);

                if (file == null)
                {
                    response.StatusCode = HttpStatusCode.BadRequest;
                    response.Content =
                        new StringContent(
                            "Ya se subió  datos para este mes, si quiere volver a subir cancele el anterior archivo.");
                    return response;
                }

                DiscountExcel contractExcel = new DiscountExcel(o.excelStream, _context, o.fileName, o.mes, o.gestion,
                    o.segmentoOrigen.ToString(), file, headerin: ExcelHeaders, sheets: 1);
                if (contractExcel.ValidateFile())
                {
                    contractExcel.toDataBase();
                    file.State = FileState.UPLOADED;
                    _context.SaveChanges();
                    response.StatusCode = HttpStatusCode.OK;
                    response.Content = new StringContent("Se subio el archivo correctamente.");
                    return response;
                }

                file.State = FileState.CANCELED;
                _context.SaveChanges();
                return contractExcel.toResponse();
            }
            catch (System.ArgumentException e)
            {
                response.StatusCode = HttpStatusCode.BadRequest;
                response.Content = new StringContent("Por favor enviar un archivo en formato excel (.xls, .xslx)" + e);
                return response;
            }
            catch (System.IO.IOException e)
            {
                response.StatusCode = HttpStatusCode.BadRequest;
                response.Content = new StringContent("El archivo es demasiado grande para ser procesado.");
                return response;
            }
            catch (System.Exception e)
            {
                response.StatusCode = HttpStatusCode.BadRequest;
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
            try
            {
                var req = await Request.Content.ReadAsMultipartAsync();
                dynamic o = HttpContentToVariables(req).Result;

                if (!((IDictionary<string, object>)o).ContainsKey("mes")
                    || !((IDictionary<string, object>)o).ContainsKey("gestion")
                    || !((IDictionary<string, object>)o).ContainsKey("segmentoOrigen")
                    || !((IDictionary<string, object>)o).ContainsKey("fileName")
                    || !((IDictionary<string, object>)o).ContainsKey("excelStream"))
                {
                    response.StatusCode = HttpStatusCode.BadRequest;
                    response.Content = new StringContent("Debe enviar mes(mm), gestion(yyyy), segmentoOrigen(id) y un archivo excel llamado file");
                    return response;
                }

                int userid = Int32.Parse(Request.Headers.GetValues("id").First());
                var file = AddFileToProcess(o.mes, o.gestion, o.segmentoOrigen, ExcelFileType.Postgrado, userid, o.fileName);

                if (file == null)
                {
                    response.StatusCode = HttpStatusCode.BadRequest;
                    response.Content = new StringContent("Ya se subió  datos para este mes, si quiere volver a subir cancele el anterior archivo.");
                    return response;
                }

                PostgradoExcel contractExcel = new PostgradoExcel(o.excelStream, _context, o.fileName, o.mes, o.gestion, o.segmentoOrigen.ToString(), file, headerin: ExcelHeaders, sheets: 1);
                if (contractExcel.ValidateFile())
                {
                    contractExcel.toDataBase();
                    file.State = FileState.UPLOADED;
                    _context.SaveChanges();
                    response.StatusCode = HttpStatusCode.OK;
                    response.Content = new StringContent("Se subio el archivo correctamente.");
                    return response;
                }
                file.State = FileState.ERROR;
                _context.SaveChanges();
                return contractExcel.toResponse();
            }
            catch (System.ArgumentException e)
            {
                response.StatusCode = HttpStatusCode.BadRequest;
                response.Content = new StringContent("Por favor enviar un archivo en formato excel (.xls, .xslx)" + e);
                return response;
            }
            catch (System.IO.IOException e)
            {
                response.StatusCode = HttpStatusCode.BadRequest;
                response.Content = new StringContent("El archivo es demasiado grande para ser procesado.");
                return response;
            }
            catch (System.Exception e)
            {
                response.StatusCode = HttpStatusCode.BadRequest;
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
            try
            {
                var req = await Request.Content.ReadAsMultipartAsync();
                dynamic o = HttpContentToVariables(req).Result;

                if (!((IDictionary<string, object>)o).ContainsKey("mes")
                    || !((IDictionary<string, object>)o).ContainsKey("gestion")
                    || !((IDictionary<string, object>)o).ContainsKey("segmentoOrigen")
                    || !((IDictionary<string, object>)o).ContainsKey("fileName")
                    || !((IDictionary<string, object>)o).ContainsKey("excelStream"))
                {
                    response.StatusCode = HttpStatusCode.BadRequest;
                    response.Content = new StringContent("Debe enviar mes(mm), gestion(yyyy), segmentoOrigen(id) y un archivo excel llamado file");
                    return response;
                }

                int userid = Int32.Parse(Request.Headers.GetValues("id").First());
                var file = AddFileToProcess(o.mes, o.gestion, o.segmentoOrigen, ExcelFileType.Pregrado, userid, o.fileName);

                if (file == null)
                {
                    response.StatusCode = HttpStatusCode.BadRequest;
                    response.Content = new StringContent("Ya se subió datos para este mes, si quiere volver a subir cancele el anterior archivo.");
                    return response;
                }

                PregradoExcel contractExcel = new PregradoExcel(o.excelStream, _context, o.fileName, o.mes, o.gestion, o.segmentoOrigen.ToString(), file, headerin: ExcelHeaders, sheets: 1);
                if (contractExcel.ValidateFile())
                {
                    contractExcel.toDataBase();
                    file.State = FileState.UPLOADED;
                    _context.SaveChanges();
                    response.StatusCode = HttpStatusCode.OK;
                    response.Content = new StringContent("Se subio el archivo correctamente.");
                    return response;
                }
                file.State = FileState.ERROR;
                _context.SaveChanges();
                return contractExcel.toResponse();
            }
            catch (System.ArgumentException e)
            {
                response.StatusCode = HttpStatusCode.BadRequest;
                response.Content = new StringContent("Por favor enviar un archivo en formato excel (.xls, .xslx)" + e);
                return response;
            }
            catch (System.IO.IOException e)
            {
                response.StatusCode = HttpStatusCode.BadRequest;
                response.Content = new StringContent("El archivo es demasiado grande para ser procesado.");
                return response;
            }
            catch (System.Exception e)
            {
                response.StatusCode = HttpStatusCode.BadRequest;
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
            try
            {
                var req = await Request.Content.ReadAsMultipartAsync();
                dynamic o = HttpContentToVariables(req).Result;

                if (!((IDictionary<string, object>)o).ContainsKey("mes")
                    || !((IDictionary<string, object>)o).ContainsKey("gestion")
                    || !((IDictionary<string, object>)o).ContainsKey("segmentoOrigen")
                    || !((IDictionary<string, object>)o).ContainsKey("fileName")
                    || !((IDictionary<string, object>)o).ContainsKey("excelStream"))
                {
                    response.StatusCode = HttpStatusCode.BadRequest;
                    response.Content = new StringContent("Debe enviar mes(mm), gestion(yyyy), segmentoOrigen(id) y un archivo excel llamado file");
                    return response;
                }

                int userid = Int32.Parse(Request.Headers.GetValues("id").First());
                var file = AddFileToProcess(o.mes, o.gestion, o.segmentoOrigen, ExcelFileType.OR, userid, o.fileName);

                if (file == null)
                {
                    response.StatusCode = HttpStatusCode.BadRequest;
                    response.Content = new StringContent("Ya se subió  datos para este mes, si quiere volver a subir cancele el anterior archivo.");
                    return response;
                }

                ORExcel contractExcel = new ORExcel(o.excelStream, _context, o.fileName, o.mes, o.gestion, o.segmentoOrigen.ToString(), file, headerin: ExcelHeaders, sheets: 1);
                if (contractExcel.ValidateFile())
                {
                    contractExcel.toDataBase();
                    file.State = FileState.UPLOADED;
                    _context.SaveChanges();
                    response.StatusCode = HttpStatusCode.OK;
                    response.Content = new StringContent("Se subio el archivo correctamente.");
                    return response;
                }
                file.State = FileState.ERROR;
                _context.SaveChanges();
                return contractExcel.toResponse();
            }
            catch (System.ArgumentException e)
            {
                response.StatusCode = HttpStatusCode.BadRequest;
                response.Content = new StringContent("Por favor enviar un archivo en formato excel (.xls, .xslx)" + e);
                return response;
            }
            catch (System.IO.IOException e)
            {
                response.StatusCode = HttpStatusCode.BadRequest;
                response.Content = new StringContent("El archivo es demasiado grande para ser procesado.");
                return response;
            }
            catch (System.Exception e)
            {
                response.StatusCode = HttpStatusCode.BadRequest;
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
            var userid = Int32.Parse(Request.Headers.GetValues("id").First());
            var user = _context.CustomUsers.FirstOrDefault(u => u.Id == userid);
            var processes = _context.DistProcesses.Include(p=>p.Branches).
                Where(p=>p.BranchesId==user.BranchesId && p.State!=ProcessState.CANCELED).
                Select(p=> new
            {
                p.BranchesId,
                p.Branches.Name,
                p.State,
                p.Id,
                p.gestion,
                p.mes
            });
            return Ok(processes);
        }

        [HttpDelete]
        [Route("api/payroll/Process/{id}")]
        public IHttpActionResult Process(int id)
        {
            var processInDB = _context.DistProcesses.FirstOrDefault(p => p.Id == id && (p.State == ProcessState.STARTED || p.State == ProcessState.ERROR || p.State == ProcessState.VALIDATED));
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
            IEnumerable<Distribution> dist = _context.Database.SqlQuery<Distribution>("SELECT a.\"Document\",a.\"TipoEmpleado\",a.\"Dependency\",a.\"PEI\","+
            " a.\"PlanEstudios\",a.\"Paralelo\",a.\"Periodo\",a.\"Project\",a.\"BussinesPartner\","+
            " a.\"Monto\",a.\"Porcentaje\",a.\"MontoDividido\",a.\"segmentoOrigen\","+
            " b.\"mes\",b.\"gestion\",e.\"Name\" as Branches ,d.\"Concept\",d.\"Name\" as CuentasContables,d.\"Indicator\" " +
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
            ex.Worksheets.Add(d.CreateDataTable(dist), "Result");


            HttpResponseMessage response = new HttpResponseMessage();
            var ms = new MemoryStream();
            ex.SaveAs(ms);
            response.StatusCode = HttpStatusCode.OK;
            response.Content = new StreamContent(ms);
            response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");
            response.Content.Headers.ContentDisposition.FileName = "Distribucion-"+id+".xlsx";
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
