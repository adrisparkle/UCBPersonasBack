using System;
using System.Collections.Generic;
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

namespace UcbBack.Controllers
{
    public class PayrollController : ApiController
    {
        private ApplicationDbContext _context;

        private  struct ErrorState
        {
            public static string PENDING = "PENDING";
            public static string REVIEWED = "REVIEWED";
        }
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
                                                             && f.BranchesId == branchid);
            if (process == null)
                return NotFound();
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
        public Dist_File AddFileToProcess(string mes, string gestion, int BranchesId,int FileType,int userid,string fileName)
        {
            var processInDB =
                _context.DistProcesses.FirstOrDefault(p =>
                    p.BranchesId == BranchesId && p.gestion == gestion && p.mes == mes);
            if (processInDB != null)
            {
                if (processInDB.State == ProcessState.STARTED || processInDB.State == ProcessState.ERROR)
                {     
                    var fileInDB = _context.FileDbs.FirstOrDefault(f => f.DistProcessId == processInDB.Id && f.DistFileTypeId == FileType && f.State == FileState.UPLOADED);
                    if (fileInDB == null)
                    {
                        processInDB.State = ProcessState.STARTED;
                        var file = new Dist_File();
                        file.Id = _context.Database.SqlQuery<int>("SELECT \"rrhh_Dist_File_sqs\".nextval FROM DUMMY;").ToList()[0];
                        file.UploadedDate = DateTime.Now;
                        file.DistFileTypeId = FileType;
                        file.Name = fileName;
                        file.State = FileState.SENDED;
                        file.CustomUserId = userid;
                        file.DistProcessId = processInDB.Id;
                        _context.FileDbs.Add(file);
                        _context.SaveChanges();
                        return file;
                    }
                    
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
                file.DistFileTypeId = FileType;
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
        public async Task<System.Dynamic.ExpandoObject> HttpContentToVariables(MultipartMemoryStreamProvider req)
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
            PayrollExcel contractExcel = new PayrollExcel(fileName:"Planilla.xlsx",headerin: 3);
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
                                     && f.DistFileTypeId == 1
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
                var file = AddFileToProcess(o.mes.ToString(), o.gestion.ToString(), o.segmentoOrigen, 1, userid, o.fileName.ToString());

                if (file == null)
                {
                    response.StatusCode = HttpStatusCode.BadRequest;
                    response.Content = new StringContent("Ya se subió  datos para este mes, si quiere volver a subir cancele el anterior archivo.");
                    return response;
                }

                PayrollExcel contractExcel = new PayrollExcel(o.excelStream, _context, o.fileName, o.mes, o.gestion, o.segmentoOrigen.ToString(),file, headerin: 3, sheets: 1);
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
        }


        [HttpGet]
        [Route("api/payroll/AcademicExcel")]
        public HttpResponseMessage GetAcademicExcel()
        {
            AcademicExcel contractExcel = new AcademicExcel(fileName: "Academico.xlsx", headerin: 3);
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
                                     && f.DistFileTypeId == 2
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
                var file = AddFileToProcess(o.mes, o.gestion, o.segmentoOrigen, 2, userid, o.fileName);

                if (file == null)
                {
                    response.StatusCode = HttpStatusCode.BadRequest;
                    response.Content = new StringContent("Ya se subió  datos para este mes, si quiere volver a subir cancele el anterior archivo.");
                    return response;
                }

                AcademicExcel contractExcel = new AcademicExcel(o.excelStream, _context, o.fileName, o.mes, o.gestion, o.segmentoOrigen.ToString(),file, headerin: 3, sheets: 1);
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
        }

        [HttpGet]
        [Route("api/payroll/DiscountExcel")]
        public HttpResponseMessage GetDiscountExcel()
        {
            DiscountExcel contractExcel = new DiscountExcel(fileName: "Descuentos.xlsx", headerin: 3);
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
                                     && f.DistFileTypeId == 3
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
                var file = AddFileToProcess(o.mes, o.gestion, o.segmentoOrigen, 3, userid, o.fileName);

                if (file == null)
                {
                    response.StatusCode = HttpStatusCode.BadRequest;
                    response.Content = new StringContent("Ya se subió  datos para este mes, si quiere volver a subir cancele el anterior archivo.");
                    return response;
                }

                DiscountExcel contractExcel = new DiscountExcel(o.excelStream, _context, o.fileName, o.mes, o.gestion, o.segmentoOrigen.ToString(), file, headerin: 3, sheets: 1);
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
        }

        [HttpGet]
        [Route("api/payroll/PostgradoExcel")]
        public HttpResponseMessage GetPostgradoExcel()
        {
            PostgradoExcel contractExcel = new PostgradoExcel(fileName: "PosGrado.xlsx", headerin: 3);
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
                                     && f.DistFileTypeId == 4
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
                var file = AddFileToProcess(o.mes, o.gestion, o.segmentoOrigen, 4, userid, o.fileName);

                if (file == null)
                {
                    response.StatusCode = HttpStatusCode.BadRequest;
                    response.Content = new StringContent("Ya se subió  datos para este mes, si quiere volver a subir cancele el anterior archivo.");
                    return response;
                }

                PostgradoExcel contractExcel = new PostgradoExcel(o.excelStream, _context, o.fileName, o.mes, o.gestion, o.segmentoOrigen.ToString(), file, headerin: 3, sheets: 1);
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
        }

        [HttpGet]
        [Route("api/payroll/PregradoExcel")]
        public HttpResponseMessage GetPregradoExcel()
        {
            PregradoExcel contractExcel = new PregradoExcel(fileName: "Pregrado.xlsx", headerin: 3);
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
                                     && f.DistFileTypeId == 5
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
                var file = AddFileToProcess(o.mes, o.gestion, o.segmentoOrigen, 5, userid, o.fileName);

                if (file == null)
                {
                    response.StatusCode = HttpStatusCode.BadRequest;
                    response.Content = new StringContent("Ya se subió  datos para este mes, si quiere volver a subir cancele el anterior archivo.");
                    return response;
                }

                PregradoExcel contractExcel = new PregradoExcel(o.excelStream, _context, o.fileName, o.mes, o.gestion, o.segmentoOrigen.ToString(), file, headerin: 3, sheets: 1);
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
        }

        [HttpGet]
        [Route("api/payroll/ORExcel")]
        public HttpResponseMessage GetORExcel()
        {
            ORExcel contractExcel = new ORExcel(fileName: "OtrasRegionales.xlsx", headerin: 3);
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
                                     && f.DistFileTypeId == 6
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
                var file = AddFileToProcess(o.mes, o.gestion, o.segmentoOrigen, 6, userid, o.fileName);

                if (file == null)
                {
                    response.StatusCode = HttpStatusCode.BadRequest;
                    response.Content = new StringContent("Ya se subió  datos para este mes, si quiere volver a subir cancele el anterior archivo.");
                    return response;
                }

                ORExcel contractExcel = new ORExcel(o.excelStream, _context, o.fileName, o.mes, o.gestion, o.segmentoOrigen.ToString(), file, headerin: 3, sheets: 1);
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
        }

        [HttpGet]
        [Route("api/payroll/GetErrors/{id}")]
        public IHttpActionResult GetErrors(int id)
        {
            var process = _context.DistProcesses.FirstOrDefault(p => p.Id == id);
            
            if (process == null)
                return NotFound();

            if (process.State != ProcessState.ERROR)
                return Ok();

            var err = _context.DistLogErroreses.Where(e => e.DistProcessId == process.Id).Include(e=>e.Error).Select(e=>new{e.Id,e.ErrorId,e.Error.Name,e.Error.Description,e.Error.Type,e.Archivos,e.CUNI});
            return Ok(err);
        }


        [HttpPost]
        [Route("api/payroll/Validate")]
        public IHttpActionResult Validate([FromBody] JObject data)
        {
            int DistProcessId = 0;
            if (!Int32.TryParse(data["DistProcessId"].ToString(), out DistProcessId))
            {
                return BadRequest();
            }
            var process = _context.DistProcesses.FirstOrDefault(p => p.Id == DistProcessId && (p.State == ProcessState.STARTED || p.State == ProcessState.ERROR));
            if (process == null)
                return NotFound();
            int userid = Int32.Parse(Request.Headers.GetValues("id").First());

            _context.Database.ExecuteSqlCommand("CALL ADMNALRRHH.VALIDATE_CUADRARDESCUENTOS(" + userid + "," + process.Id + ")");
            _context.Database.ExecuteSqlCommand("CALL ADMNALRRHH.VALIDATE_ACADSUM(" + userid + "," + process.Id + ")");
            _context.Database.ExecuteSqlCommand("CALL ADMNALRRHH.VALIDATE_CE(" + userid + "," + process.Id + ")");
            _context.Database.ExecuteSqlCommand("CALL ADMNALRRHH.VALIDATE_OD(" + userid + "," + process.Id + ")");
            _context.Database.ExecuteSqlCommand("CALL ADMNALRRHH.VALIDATE_OR(" + userid + "," + process.Id + ")");
            _context.Database.ExecuteSqlCommand("CALL ADMNALRRHH.VALIDATE_OTHERINCOMES(" + userid + "," + process.Id + ")");
            _context.Database.ExecuteSqlCommand("CALL ADMNALRRHH.VALIDATE_TIPOEMPLEADO(" + userid + "," + process.Id + ")");


            var err = _context.DistLogErroreses.Where(e => e.DistProcessId == process.Id).Include(e => e.Error).Select(e => new { e.Id, e.ErrorId, e.Error.Name, e.Error.Description,e.Error.Type, e.Archivos, e.CUNI });
            if (err.Count()>0)
            {
                
                if(err.Select(e=>e.Type=="E").Count()>0)
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
        [Route("api/payroll/Distribute/{id}")]
        public IHttpActionResult Distribute([FromBody] int id)
        {
            var process = _context.DistProcesses.FirstOrDefault(p => p.Id == id && (p.State == ProcessState.VALIDATED || p.State == ProcessState.WARNING));
            if (process == null)
                return NotFound();

            _context.Database.ExecuteSqlCommand("CALL ADMNALRRHH.SET_PERCENT(" + process.Id + ")");
            _context.Database.ExecuteSqlCommand("CALL ADMNALRRHH.DIST_PERCENTS(" + process.Id + ")");
            _context.Database.ExecuteSqlCommand("CALL ADMNALRRHH.DIST_COSTS(" + process.Id + ")");
            process.State = ProcessState.PROCESSED;
            _context.SaveChanges();
            return Ok("Se procesó la información");
        }

        [HttpDelete]
        [Route("api/payroll/Process/{id}")]
        public IHttpActionResult Process(int id)
        {
            var processInDB = _context.DistProcesses.FirstOrDefault(p => p.Id == id && (p.State == ProcessState.STARTED || p.State == ProcessState.ERROR));
            if (processInDB == null)
                return NotFound();
            processInDB.State = ProcessState.CANCELED;
            _context.SaveChanges();
            return Ok("Proceso Cancelado");
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

            var sum = _context.DistPayrolls.Where(p => p.DistFileId == file.Id).ToList();

            dynamic res = new JObject();
            res.BasicSalary = sum.Sum(p => p.BasicSalary);
            res.AntiquityBonus = sum.Sum(p => p.AntiquityBonus);
            res.OtherIncome = sum.Sum(p => p.OtherIncome);
            res.TeachingIncome = sum.Sum(p => p.TeachingIncome);
            res.OtherAcademicIncomes = sum.Sum(p => p.OtherAcademicIncomes);
            res.Reintegro = sum.Sum(p => p.Reintegro);
            res.TotalAmountEarned = sum.Sum(p => p.TotalAmountEarned);
            res.AFPLaboral = sum.Sum(p => p.AFPLaboral);
            res.RcIva = sum.Sum(p => p.RcIva);
            res.Discounts = sum.Sum(p => p.Discounts);
            res.TotalAmountDiscounts = sum.Sum(p => p.TotalAmountDiscounts);
            res.TotalAfterDiscounts = sum.Sum(p => p.TotalAfterDiscounts);
            res.AFPPatronal = sum.Sum(p => p.AFPPatronal);
            res.SeguridadCortoPlazoPatronal = sum.Sum(p => p.SeguridadCortoPlazoPatronal);
            res.ProvAguinaldo = sum.Sum(p => p.ProvAguinaldo);
            res.ProvPrimas = sum.Sum(p => p.ProvPrimas);
            res.ProvIndeminizacion = sum.Sum(p => p.ProvIndeminizacion);

            return Ok(res);
        }

    }



}
