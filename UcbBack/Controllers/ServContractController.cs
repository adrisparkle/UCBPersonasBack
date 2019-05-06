using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web.Http;
using ClosedXML.Excel;
using Newtonsoft.Json.Linq;
using UcbBack.Logic;
using UcbBack.Models;
using UcbBack.Models.Not_Mapped;
using UcbBack.Models.Not_Mapped.CustomDataAnnotations;
using System.Data;
using System.Threading.Tasks;
using Sap.Data.Hana;
using UcbBack.Logic.ExcelFiles;
using UcbBack.Logic.ExcelFiles.Serv;
using UcbBack.Models.Not_Mapped.ViewMoldes;
using UcbBack.Models.Serv;

namespace UcbBack.Controllers
{
    public class ServContractController : ApiController
    {
        private ApplicationDbContext _context;
        private ValidateAuth auth;


        public ServContractController()
        {
            _context = new ApplicationDbContext();
            auth = new ValidateAuth();
        }

        public IHttpActionResult Get(int id)
        {
            var user = auth.getUser(Request);
            var query = "select * from " + CustomSchema.Schema + ".\"Serv_Process\" " +
                        " order by (" +
                        "   case when \"State\" = " + ServProcess.Serv_FileState.PendingApproval + " then 1 " +
                        " when \"State\" = " + ServProcess.Serv_FileState.Started + " then 2 " +
                        " when \"State\" = " + ServProcess.Serv_FileState.INSAP + " then 3 " +
                        " when \"State\" = " + ServProcess.Serv_FileState.Canceled + " then 4 " +
                        " when \"State\" = " + ServProcess.Serv_FileState.ERROR + " then 5 " +
                        " end) asc, " +
                        " \"CreatedAt\" desc;";
            var rawresult = _context.Database.SqlQuery<Civil>(query).ToList();

            if (rawresult.Count() == 0)
                return NotFound();

            var res = auth.filerByRegional(rawresult.AsQueryable(), user);

            if (res.Count() == 0)
                return Unauthorized();

            return Ok(res.FirstOrDefault());
        }

        [HttpPost]
        [Route("api/ServContractgenerateExcel/")]
        public HttpResponseMessage generateExcel(JObject data)
        {
            HttpResponseMessage response = new HttpResponseMessage();
            var list = data["list"].ToObject<List<int>>();
            var ex = new XLWorkbook();

            var excelData = getData(list, data["tipo"].ToString());

            ex.Worksheets.Add(excelData, "Plantilla_" + data["tipo"].ToString());
            var ms = new MemoryStream();
            ex.SaveAs(ms);
            response.StatusCode = HttpStatusCode.OK;
            response.Content = new StreamContent(ms);
            response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");
            response.Content.Headers.ContentDisposition.FileName = "Plantilla_" + data["tipo"].ToString() + ".xlsx";
            response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            response.Content.Headers.ContentLength = ms.Length;
            ms.Seek(0, SeekOrigin.Begin);
            return response;
        }

        [HttpDelete]
        [Route("api/ServContract/UploadFile")]
        public IHttpActionResult DeleteProcess(JObject data)
        {
            var user = auth.getUser(Request);
            //todo add validation of user by branch
            int branchesid;
            if (data["BranchesId"] == null || data["FileType"] == null || !Int32.TryParse(data["BranchesId"].ToString(), out branchesid))
            {
                ModelState.AddModelError("Mal Formato", "Debes enviar BranchesId y FileType");
                return BadRequest();
            }

            string type = data["FileType"].ToString();
            var process = _context.ServProcesses.FirstOrDefault(x =>
                x.BranchesId == branchesid && x.FileType == type && x.State == ServProcess.Serv_FileState.Started);
            if (process == null)
                return NotFound();
            process.State = ServProcess.Serv_FileState.Canceled;
            process.LastUpdatedBy = user.Id;
            _context.SaveChanges();
            return Ok();
        }
        [HttpPost]
        [Route("api/ServContract/UploadFile")]
        public async Task<HttpResponseMessage> UploadORExcel()
        {
            var response = new HttpResponseMessage();
            try
            {
                var req = await Request.Content.ReadAsMultipartAsync();
                dynamic o = HttpContentToVariables(req).Result;

                if (!((IDictionary<string, object>)o).ContainsKey("BranchesId")
                    || !((IDictionary<string, object>)o).ContainsKey("FileType")
                    || !((IDictionary<string, object>)o).ContainsKey("fileName")
                    || !((IDictionary<string, object>)o).ContainsKey("excelStream")
                    || !o.fileName.ToString().EndsWith(".xlsx"))
                {
                    response.StatusCode = HttpStatusCode.BadRequest;
                    response.Headers.Add("UploadErrors", "{ \"Faltan datos\": \"Debe enviar mes(mm), gestion(yyyy), segmentoOrigen(id) y un archivo excel llamado file (en formato .xlsx)\"}");
                    response.Content = new StringContent("Debe enviar mes(mm), gestion(yyyy), segmentoOrigen(id) y un archivo excel llamado file");
                    return response;
                }


                //todo validate FileType
                // ...


                string realFileName;
                if (!verifyName(o.fileName, o.BranchesId, o.FileType, out realFileName))
                {
                    response.StatusCode = HttpStatusCode.BadRequest;
                    response.Headers.Add("UploadErrors", "{ \"Nombre Incorrecto\": \"El archivo enviado no cumple con la regla de nombres. Nombre sugerido: " + realFileName + "\"}");
                    response.Content = new StringContent("El archivo enviado no cumple con la regla de nombres.");
                    return response;
                }

                int userid = Int32.Parse(Request.Headers.GetValues("id").First());
                ServProcess file = AddFileToProcess(Int32.Parse(o.BranchesId.ToString()), o.FileType.ToString(), userid);

                if (file == null)
                {
                    response.StatusCode = HttpStatusCode.BadRequest;
                    response.Headers.Add("UploadErrors", "{ \"Ya se Subio archivos para este mes\": \"Ya se subio  datos para este mes, si quiere volver a subir cancele el anterior archivo.\"}");
                    response.Content = new StringContent("Ya se subió  datos para este mes, si quiere volver a subir cancele el anterior archivo.");
                    return response;
                }

                DynamicExcelToDB(o.FileType,o,file,out response);
                return response;
            }
            catch (System.ArgumentException e)
            {
                response.StatusCode = HttpStatusCode.BadRequest;
                response.Headers.Add("UploadErrors", "{ \"Formato Archivo Invalido\": \"Por favor enviar un archivo en formato excel (.xlsx)\"}");
                response.Content = new StringContent("Por favor enviar un archivo en formato excel (.xlsx)" + e);
                return response;
            }
            catch (System.IO.IOException e)
            {
                Console.WriteLine(e);
                response.StatusCode = HttpStatusCode.BadRequest;
                response.Headers.Add("UploadErrors", "{ \"Archivo demasiado grande\": \"El archivo es demasiado grande para ser procesado.\"}");
                response.Content = new StringContent("El archivo es demasiado grande para ser procesado.");
                return response;
            }
            catch (HanaException e)
            {
                if (e.NativeError == 258)
                {
                    Console.WriteLine(e);
                    response.StatusCode = HttpStatusCode.BadRequest;
                    response.Headers.Add("Error en conexion con SAP", "{ \"La conexion con SAP se perdio\": \"No se pudo validar el archivo con con SAP.\"}");
                    response.Content = new StringContent("Error conexion SAP");
                    return response;
                }
                Console.WriteLine(e);
                response.StatusCode = HttpStatusCode.BadRequest;
                response.Headers.Add("Error en conexion con SAP", "{ \"La conexion con SAP se perdio\": \"No se pudo validar el archivo con con SAP.\"}");
                response.Content = new StringContent("Error conexion SAP");
                return response;
            }
            /*catch (System.Exception e)
            {
                Console.WriteLine(e);
                response.StatusCode = HttpStatusCode.BadRequest;
                response.Headers.Add("UploadErrors", "{ \"Existen Enlaces a otros archivos\": \"Existen celdas con referencias a otros archivos.\"}");
                response.Content = new StringContent("Por favor enviar un archivo en formato excel sin referencias a otros libros excel o formulas(.xls, .xslx)");
                return response;
            }*/
        }

        [HttpGet]
        [Route("api/ServContract/GetDetail/{id}")]
        public IHttpActionResult GetDetail(int id)
        {
            var process = _context.ServProcesses.FirstOrDefault(p => p.Id == id);
            
            if (process == null)
                return NotFound();
            switch (process.FileType)
            {
                case ServProcess.Serv_FileType.Varios:
                    return Ok(_context.ServVarioses.Where(x => x.Serv_ProcessId == process.Id));
                case ServProcess.Serv_FileType.Carrera:
                    return Ok(_context.ServCarreras.Where(x => x.Serv_ProcessId == process.Id));

                case ServProcess.Serv_FileType.Paralelo:
                    return Ok(_context.ServParalelos.Where(x => x.Serv_ProcessId == process.Id));

                case ServProcess.Serv_FileType.Proyectos:
                    return Ok(_context.ServProyectoses.Where(x => x.Serv_ProcessId == process.Id));

            }
            return Ok();
        }

        [HttpPost]
        [Route("api/ServContract/CheckUpload")]
        public IHttpActionResult CheckUpload([FromBody] JObject upload)
        {
            int branchid = 0;
            if (upload["FileType"] == null || upload["BranchesId"] == null || !Int32.TryParse(upload["BranchesId"].ToString(), out branchid))
                return BadRequest("Debes enviar Tipo de Archivo y segmentoOrigen");

            var FileType = upload["FileType"].ToString();

            var process = _context.ServProcesses.FirstOrDefault(f => f.BranchesId == branchid
                                                                     && f.State == ServProcess.Serv_FileState.Started
                                                                     && f.FileType == FileType);
            if (process == null)
                return Ok();

            List<string> tipos = new List<string>();
            tipos.Add(FileType);

            dynamic res = new JObject();
            res.array = JToken.FromObject(tipos);
            res.id = process.Id;
            res.state = process.State;
            return Ok(res);
        }

        [NonAction]
        private DataTable getData(List<int> list,string type)
        {
            var d = new Distribution();

            switch (type)
            {
                case ServProcess.Serv_FileType.Varios:
                    var res = (from bp in _context.Civils
                        where list.Contains(bp.Id)
                        select new Serv_VariosViewModel()
                        {
                            Codigo_Socio = bp.SAPId,
                            Nombre_Socio = bp.FullName,
                            Cod_Dependencia = "",
                            PEI_PO = "",
                            Nombre_del_Servicio = "",
                            Objeto_del_Contrato = "",
                            Cuenta_Asignada = "",
                            Monto_Contrato = 0,
                            Monto_IUE = 0,
                            Monto_IT = 0,
                            Monto_a_Pagar = 0,
                            Observaciones = "",
                        }).ToList();
                    return d.CreateDataTable(res);
                case ServProcess.Serv_FileType.Carrera:
                    var res1 = (from bp in _context.Civils
                        where list.Contains(bp.Id)
                        select new Serv_PregradoViewModel()
                        {
                            Codigo_Socio = bp.SAPId,
                            Nombre_Socio = bp.FullName,
                            Cod_Dependencia = "",
                            PEI_PO = "",
                            Nombre_del_Servicio = "",
                            Codigo_Carrera = "",
                            Documento_Base = "",
                            Postulante = "",
                            Tipo_Tarea_Asignada= "",
                            Cuenta_Asignada = "",
                            Monto_Contrato = 0,
                            Monto_IUE = 0,
                            Monto_IT = 0,
                            Monto_a_Pagar = 0,
                            Observaciones = "",
                        }).ToList();
                    return d.CreateDataTable(res1);
                case ServProcess.Serv_FileType.Paralelo:
                    var res2 = (from bp in _context.Civils
                        where list.Contains(bp.Id)
                        select new Serv_ReemplazoViewModel()
                        {
                            Codigo_Socio = bp.SAPId,
                            Nombre_Socio = bp.FullName,
                            Cod_Dependencia = "",
                            PEI_PO = "",
                            Nombre_del_Servicio = "",
                            Periodo_Academico = "",
                            Sigla_Asignatura = "",
                            Paralelo = "",
                            Código_Paralelo_SAP = "",
                            Cuenta_Asignada = "",
                            Monto_Contrato = 0,
                            Monto_IUE = 0,
                            Monto_IT = 0,
                            Monto_a_Pagar = 0,
                            Observaciones = "",
                        }).ToList();
                    return d.CreateDataTable(res2);
                case ServProcess.Serv_FileType.Proyectos:
                    var res3 = (from bp in _context.Civils
                        where list.Contains(bp.Id)
                        select new Serv_ProyectosViewModel()
                        {
                            Codigo_Socio = bp.SAPId,
                            Nombre_Socio = bp.FullName,
                            Cod_Dependencia = "",
                            PEI_PO = "",
                            Nombre_del_Servicio = "",
                            Código_Proyecto_SAP = "",
                            Nombre_del_Proyecto = "",
                            Versión = "",
                            Periodo_Académico = "",
                            Tipo_de_Tarea_Asignada = "",
                            Cuenta_Asignada = "",
                            Monto_Contrato = 0,
                            Monto_IUE = 0,
                            Monto_IT = 0,
                            Monto_a_Pagar = 0,
                            Observaciones = "",
                        }).ToList();
                    return d.CreateDataTable(res3);
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
                if (varname == "\"BranchesId\"")
                {
                    res.BranchesId = Int32.Parse(contentPart.ReadAsStringAsync().Result.ToString());
                }
                else if (varname == "\"FileType\"")
                {
                    res.FileType = contentPart.ReadAsStringAsync().Result.ToString();
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

        [NonAction]
        private bool verifyName(string fileName, int branchId, string fileType,
            out string realfileName)
        {
            string Abr = _context.Branch.Where(x => x.Id == branchId).Select(x => x.Abr).FirstOrDefault();
            realfileName = Abr + "-" + fileType;
            return fileName.Split('.')[0].Equals(realfileName);
        }

        [NonAction]
        private ServProcess AddFileToProcess(int BranchesId, string FileType, int userid)
        {
            var processInDB = _context.ServProcesses.FirstOrDefault(p =>
                    p.BranchesId == BranchesId && p.FileType == FileType && p.State == ServProcess.Serv_FileState.Started);

            //if exist a process of the same type, cancel and create a new one
            if (processInDB != null )
            {
                processInDB.State = ServProcess.Serv_FileState.Canceled;
            }
            //create new process
            var process = new ServProcess();
            process.Id = process.GetNextId(_context);
            process.CreatedAt = DateTime.Now;
            process.FileType = FileType;
            process.State = ServProcess.Serv_FileState.Started;
            process.CreatedBy = userid;
            process.BranchesId = BranchesId;

            _context.ServProcesses.Add(process);
            _context.SaveChanges();
            return process;
        }

        [NonAction]
        private void DynamicExcelToDB(string FileType, dynamic o, ServProcess file,  out HttpResponseMessage response)
        {
            response = new HttpResponseMessage();
            switch (FileType)
            {
                case ServProcess.Serv_FileType.Varios:
                    Serv_VariosExcel ExcelFile = new Serv_VariosExcel(o.excelStream, _context, o.fileName, file,headerin:1,sheets:1);
                    if (ExcelFile.ValidateFile())
                    {
                        ExcelFile.toDataBase();
                        file.State = ServProcess.Serv_FileState.Started;
                        _context.SaveChanges();
                        response.StatusCode = HttpStatusCode.OK;
                        response.Content = new StringContent("Se subio el archivo correctamente.");
                        _context.SaveChanges();
                    }
                    else
                    {
                        file.State = ServProcess.Serv_FileState.ERROR;
                        _context.SaveChanges();
                        response = ExcelFile.toResponse();
                    }
                    break;

                case ServProcess.Serv_FileType.Carrera:
                    Serv_CarreraExcel ExcelFile2 = new Serv_CarreraExcel(o.excelStream, _context, o.fileName, file, headerin: 1, sheets: 1);
                    if (ExcelFile2.ValidateFile())
                    {
                        ExcelFile2.toDataBase();
                        file.State = ServProcess.Serv_FileState.Started;
                        _context.SaveChanges();
                        response.StatusCode = HttpStatusCode.OK;
                        response.Content = new StringContent("Se subio el archivo correctamente.");
                        _context.SaveChanges();
                    }
                    else
                    {
                        file.State = ServProcess.Serv_FileState.ERROR;
                        _context.SaveChanges();
                        response = ExcelFile2.toResponse();
                    }
                    break;

                case ServProcess.Serv_FileType.Proyectos:
                    Serv_ProyectosExcel ExcelFile3 = new Serv_ProyectosExcel(o.excelStream, _context, o.fileName, file, headerin: 1, sheets: 1);
                    if (ExcelFile3.ValidateFile())
                    {
                        ExcelFile3.toDataBase();
                        file.State = ServProcess.Serv_FileState.Started;
                        _context.SaveChanges();
                        response.StatusCode = HttpStatusCode.OK;
                        response.Content = new StringContent("Se subio el archivo correctamente.");
                        _context.SaveChanges();
                    }
                    else
                    {
                        file.State = ServProcess.Serv_FileState.ERROR;
                        _context.SaveChanges();
                        response = ExcelFile3.toResponse();
                    }
                    break;
                case ServProcess.Serv_FileType.Paralelo:
                    Serv_ParaleloExcel ExcelFile4 = new Serv_ParaleloExcel(o.excelStream, _context, o.fileName, file, headerin: 1, sheets: 1);
                    if (ExcelFile4.ValidateFile())
                    {
                        ExcelFile4.toDataBase();
                        file.State = ServProcess.Serv_FileState.Started;
                        _context.SaveChanges();
                        response.StatusCode = HttpStatusCode.OK;
                        response.Content = new StringContent("Se subio el archivo correctamente.");
                        _context.SaveChanges();
                    }
                    else
                    {
                        file.State = ServProcess.Serv_FileState.ERROR;
                        _context.SaveChanges();
                        response = ExcelFile4.toResponse();
                    }
                    break;
            }
        }

        

    }
}
