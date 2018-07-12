using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Formatting;
using System.Threading.Tasks;
using System.Web;
using System.Web.Http;
using Newtonsoft.Json.Linq;
using UcbBack.Logic.ExcelFiles;
using UcbBack.Models;

namespace UcbBack.Controllers
{
    public class PayrollController : ApiController
    {
        private ApplicationDbContext _context;

        public PayrollController()
        {
            _context = new ApplicationDbContext();
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
                    res.mes = contentPart.ReadAsStringAsync().Result;
                }
                else if (varname == "\"gestion\"")
                {
                    res.gestion = contentPart.ReadAsStringAsync().Result;
                }
                else if (varname == "\"segmentoOrigen\"")
                {
                    res.segmentoOrigen = contentPart.ReadAsStringAsync().Result;
                }
                else if (varname == "\"uploadfile\"")
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
                    response.Content = new StringContent("Debe enviar mes, gestion, segmentoOrigen y un archivo excel llamado uploadfile");
                    return response;
                }

                PayrollExcel contractExcel = new PayrollExcel(o.excelStream, _context, o.fileName, o.mes, o.gestion, o.segmentoOrigen, headerin: 3, sheets: 1);
                if (contractExcel.ValidateFile())
                {
                    contractExcel.toDataBase();
                    response.StatusCode = HttpStatusCode.OK;
                    response.Content = new StringContent("Se subio el archivo correctamente.");
                    return response;
                }
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
                    response.Content = new StringContent("Debe enviar mes, gestion y segmentoOrigen");
                    return response;
                }

                AcademicExcel contractExcel = new AcademicExcel(o.excelStream, _context, o.fileName, o.mes, o.gestion, o.segmentoOrigen, headerin: 3, sheets: 1);
                if (contractExcel.ValidateFile())
                {
                    contractExcel.toDataBase();
                    response.StatusCode = HttpStatusCode.OK;
                    response.Content = new StringContent("Se subio el archivo correctamente.");
                    return response;
                }
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
                    response.Content = new StringContent("Debe enviar mes, gestion y segmentoOrigen");
                    return response;
                }

                DiscountExcel contractExcel = new DiscountExcel(o.excelStream, _context, o.fileName, o.mes, o.gestion, o.segmentoOrigen, headerin: 3, sheets: 1);
                if (contractExcel.ValidateFile())
                {
                    contractExcel.toDataBase();
                    response.StatusCode = HttpStatusCode.OK;
                    response.Content = new StringContent("Se subio el archivo correctamente.");
                    return response;
                }
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
                    response.Content = new StringContent("Debe enviar mes, gestion y segmentoOrigen");
                    return response;
                }

                PostgradoExcel contractExcel = new PostgradoExcel(o.excelStream, _context, o.fileName, o.mes, o.gestion, o.segmentoOrigen, headerin: 3, sheets: 1);
                if (contractExcel.ValidateFile())
                {
                    contractExcel.toDataBase();
                    response.StatusCode = HttpStatusCode.OK;
                    response.Content = new StringContent("Se subio el archivo correctamente.");
                    return response;
                }
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
                    response.Content = new StringContent("Debe enviar mes, gestion y segmentoOrigen");
                    return response;
                }

                PregradoExcel contractExcel = new PregradoExcel(o.excelStream, _context, o.fileName, o.mes, o.gestion, o.segmentoOrigen, headerin: 3, sheets: 1);
                if (contractExcel.ValidateFile())
                {
                    contractExcel.toDataBase();
                    response.StatusCode = HttpStatusCode.OK;
                    response.Content = new StringContent("Se subio el archivo correctamente.");
                    return response;
                }
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
                    response.Content = new StringContent("Debe enviar mes, gestion y segmentoOrigen");
                    return response;
                }

                ORExcel contractExcel = new ORExcel(o.excelStream, _context, o.fileName, o.mes, o.gestion, o.segmentoOrigen, headerin: 3, sheets: 1);
                if (contractExcel.ValidateFile())
                {
                    contractExcel.toDataBase();
                    response.StatusCode = HttpStatusCode.OK;
                    response.Content = new StringContent("Se subio el archivo correctamente.");
                    return response;
                }
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
        [Route("api/payroll/Distribute")]
        public IHttpActionResult Distribute()
        {
            var sql = _context.Database.ExecuteSqlCommand("");
            return Ok("Se Distribuyó las planillas con éxito");
        }

    }



}
