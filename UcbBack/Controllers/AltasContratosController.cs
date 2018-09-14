using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Http;
using UcbBack.Logic.ExcelFiles;
using UcbBack.Models;

namespace UcbBack.Controllers
{
    public class AltasContratosController : ApiController
    {
        private ApplicationDbContext _context;

        [NonAction]
        private async Task<System.Dynamic.ExpandoObject> HttpContentToVariables(MultipartMemoryStreamProvider req)
        {
            dynamic res = new System.Dynamic.ExpandoObject();
            foreach (HttpContent contentPart in req.Contents)
            {
                var contentDisposition = contentPart.Headers.ContentDisposition;
                string varname = contentDisposition.Name;
                if (varname == "\"segmentoOrigen\"")
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

        [HttpPost]
        [Route("api/Contratos/AltasExcel")]
        public async Task<HttpResponseMessage> UploadPayrollExcel()
        {
            var response = new HttpResponseMessage();
            AltasPersonalExcel ExcelFile = null;
            try
            {
                var req = await Request.Content.ReadAsMultipartAsync();
                dynamic o = HttpContentToVariables(req).Result;

                if (!((IDictionary<string, object>)o).ContainsKey("segmentoOrigen")
                    || !((IDictionary<string, object>)o).ContainsKey("fileName")
                    || !((IDictionary<string, object>)o).ContainsKey("excelStream")
                    || !o.fileName.ToString().EndsWith(".xlsx"))
                {
                    response.StatusCode = HttpStatusCode.BadRequest;
                    response.Headers.Add("UploadErrors", "{ \"Faltan datos\": \"Debe enviar segmentoOrigen(id) y un archivo excel llamado file (en formato .xlsx)\"}");
                    response.Content = new StringContent("segmentoOrigen(id) y un archivo excel llamado file");
                    return response;
                }

                int userid = 0;
                if (!Int32.TryParse(Request.Headers.GetValues("id").First(), out userid))
                {
                    response.StatusCode = HttpStatusCode.Unauthorized;
                    return response;
                }

                ExcelFile = new AltasPersonalExcel(o.excelStream,_context, o.fileName,o.segmentoOrigen.ToString());

                if (ExcelFile.ValidateFile())
                {
                    ExcelFile.toDataBase();
                    _context.SaveChanges();
                    response.StatusCode = HttpStatusCode.OK;
                    response.Content = new StringContent("Se subio el archivo correctamente.");
                    return response;
                }
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
    }
}
