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
    public class PayrollController : ApiController
    {
        private ApplicationDbContext _context;

        public PayrollController()
        {
            _context = new ApplicationDbContext();
        }

        [HttpPost]
        [Route("api/payroll/UploadPayrollExcel")]
        public async Task<HttpResponseMessage> UploadPayrollExcel()
        {
            var response = new HttpResponseMessage();
            try
            {
                var req = await Request.Content.ReadAsMultipartAsync();
                foreach (HttpContent contentPart in req.Contents)
                {
                    Stream stream = await contentPart.ReadAsStreamAsync();
                    var contentDisposition = contentPart.Headers.ContentDisposition;
                    string fileName = String.IsNullOrEmpty(contentDisposition.FileName) ? "" : contentDisposition.FileName.Trim('"');
                    PayrollExcel contractExcel = new PayrollExcel(stream,_context,fileName,headerin:3,sheets:1);

                    if (contractExcel.ValidateFile())
                    {
                        contractExcel.toDataBase();
                        response.StatusCode = HttpStatusCode.OK;
                        response.Content = new StringContent("Se subio el archivo correctamente.");
                        return response;
                    }
                    return contractExcel.toResponse();
                }
                response.StatusCode = HttpStatusCode.BadRequest;
                response.Content = new StringContent("Por favor enviar un archivo en formato excel (.xls, .xslx)");
                return response;
            }
            catch (System.ArgumentException e)
            {
                response.StatusCode = HttpStatusCode.BadRequest;
                response.Content = new StringContent("Por favor enviar un archivo en formato excel (.xls, .xslx)" + e);
                return response;
            }
        }


        [HttpPost]
        [Route("api/payroll/UploadAcademicExcel")]
        public async Task<HttpResponseMessage> UploadAcademicExcel()
        {
            var response = new HttpResponseMessage();
            try
            {
                var req = await Request.Content.ReadAsMultipartAsync();
                foreach (HttpContent contentPart in req.Contents)
                {
                    Stream stream = await contentPart.ReadAsStreamAsync();
                    var contentDisposition = contentPart.Headers.ContentDisposition;
                    string fileName = String.IsNullOrEmpty(contentDisposition.FileName) ? "" : contentDisposition.FileName.Trim('"');
                    AcademicExcel contractExcel = new AcademicExcel(stream, _context, fileName, headerin: 3, sheets: 1);

                    if (contractExcel.ValidateFile())
                    {
                        contractExcel.toDataBase();
                        response.StatusCode = HttpStatusCode.OK;
                        response.Content = new StringContent("Se subio el archivo correctamente.");
                        return response;
                    }
                    return contractExcel.toResponse();
                }
                response.StatusCode = HttpStatusCode.BadRequest;
                response.Content = new StringContent("Por favor enviar un archivo en formato excel (.xls, .xslx)");
                return response;
            }
            catch (System.ArgumentException e)
            {
                response.StatusCode = HttpStatusCode.BadRequest;
                response.Content = new StringContent("Por favor enviar un archivo en formato excel (.xls, .xslx)" + e);
                return response;
            }
        }

        [HttpPost]
        [Route("api/payroll/UploadDiscountExcel")]
        public async Task<HttpResponseMessage> UploadDiscountExcel()
        {
            var response = new HttpResponseMessage();
            try
            {
                var req = await Request.Content.ReadAsMultipartAsync();
                foreach (HttpContent contentPart in req.Contents)
                {
                    Stream stream = await contentPart.ReadAsStreamAsync();
                    var contentDisposition = contentPart.Headers.ContentDisposition;
                    string fileName = String.IsNullOrEmpty(contentDisposition.FileName) ? "" : contentDisposition.FileName.Trim('"');
                    DiscountExcel disExcel = new DiscountExcel(stream, _context, fileName, headerin: 3, sheets: 1);

                    if (disExcel.ValidateFile())
                    {
                        disExcel.toDataBase();
                        response.StatusCode = HttpStatusCode.OK;
                        response.Content = new StringContent("Se subio el archivo correctamente.");
                        return response;
                    }
                    return disExcel.toResponse();
                }
                response.StatusCode = HttpStatusCode.BadRequest;
                response.Content = new StringContent("Por favor enviar un archivo en formato excel (.xls, .xslx)");
                return response;
            }
            catch (System.ArgumentException e)
            {
                response.StatusCode = HttpStatusCode.BadRequest;
                response.Content = new StringContent("Por favor enviar un archivo en formato excel (.xls, .xslx)" + e);
                return response;
            }
        }

        [HttpPost]
        [Route("api/payroll/UploadPostgradoExcel")]
        public async Task<HttpResponseMessage> UploadPostgradoExcel()
        {
            var response = new HttpResponseMessage();
            try
            {
                var req = await Request.Content.ReadAsMultipartAsync();
                foreach (HttpContent contentPart in req.Contents)
                {
                    Stream stream = await contentPart.ReadAsStreamAsync();
                    var contentDisposition = contentPart.Headers.ContentDisposition;
                    string fileName = String.IsNullOrEmpty(contentDisposition.FileName) ? "" : contentDisposition.FileName.Trim('"');
                    PostgradoExcel disExcel = new PostgradoExcel(stream, _context, fileName, headerin: 3, sheets: 1);

                    if (disExcel.ValidateFile())
                    {
                        disExcel.toDataBase();
                        response.StatusCode = HttpStatusCode.OK;
                        response.Content = new StringContent("Se subio el archivo correctamente.");
                        return response;
                    }
                    return disExcel.toResponse();
                }
                response.StatusCode = HttpStatusCode.BadRequest;
                response.Content = new StringContent("Por favor enviar un archivo en formato excel (.xls, .xslx)");
                return response;
            }
            catch (System.ArgumentException e)
            {
                response.StatusCode = HttpStatusCode.BadRequest;
                response.Content = new StringContent("Por favor enviar un archivo en formato excel (.xls, .xslx)" + e);
                return response;
            }
        }


        [HttpPost]
        [Route("api/payroll/UploadPregradoExcel")]
        public async Task<HttpResponseMessage> UploadPregradoExcel()
        {
            var response = new HttpResponseMessage();
            try
            {
                var req = await Request.Content.ReadAsMultipartAsync();
                foreach (HttpContent contentPart in req.Contents)
                {
                    Stream stream = await contentPart.ReadAsStreamAsync();
                    var contentDisposition = contentPart.Headers.ContentDisposition;
                    string fileName = String.IsNullOrEmpty(contentDisposition.FileName) ? "" : contentDisposition.FileName.Trim('"');
                    PregradoExcel disExcel = new PregradoExcel(stream, _context, fileName, headerin: 3, sheets: 1);

                    if (disExcel.ValidateFile())
                    {
                        disExcel.toDataBase();
                        response.StatusCode = HttpStatusCode.OK;
                        response.Content = new StringContent("Se subio el archivo correctamente.");
                        return response;
                    }
                    return disExcel.toResponse();
                }
                response.StatusCode = HttpStatusCode.BadRequest;
                response.Content = new StringContent("Por favor enviar un archivo en formato excel (.xls, .xslx)");
                return response;
            }
            catch (System.ArgumentException e)
            {
                response.StatusCode = HttpStatusCode.BadRequest;
                response.Content = new StringContent("Por favor enviar un archivo en formato excel (.xls, .xslx)" + e);
                return response;
            }
        }

        [HttpPost]
        [Route("api/payroll/UploadORExcel")]
        public async Task<HttpResponseMessage> UploadORExcel()
        {
            var response = new HttpResponseMessage();
            try
            {
                var req = await Request.Content.ReadAsMultipartAsync();
                foreach (HttpContent contentPart in req.Contents)
                {
                    Stream stream = await contentPart.ReadAsStreamAsync();
                    var contentDisposition = contentPart.Headers.ContentDisposition;
                    string fileName = String.IsNullOrEmpty(contentDisposition.FileName) ? "" : contentDisposition.FileName.Trim('"');
                    ORExcel disExcel = new ORExcel(stream, _context, fileName, headerin: 3, sheets: 1);

                    if (disExcel.ValidateFile())
                    {
                        disExcel.toDataBase();
                        response.StatusCode = HttpStatusCode.OK;
                        response.Headers.Add("UploadErrors", "{}");
                        response.Content = new StringContent("Se subio el archivo correctamente.");
                        return response;
                    }
                    return disExcel.toResponse();
                }
                response.StatusCode = HttpStatusCode.BadRequest;
                response.Content = new StringContent("Por favor enviar un archivo en formato excel (.xls, .xslx)");
                return response;
            }
            catch (System.ArgumentException e)
            {
                response.StatusCode = HttpStatusCode.BadRequest;
                response.Content = new StringContent("Por favor enviar un archivo en formato excel (.xls, .xslx)" + e);
                return response;
            }
        }

    }



}
