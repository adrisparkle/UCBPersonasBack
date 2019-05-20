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
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Sap.Data.Hana;
using UcbBack.Logic.ExcelFiles;
using UcbBack.Logic.ExcelFiles.Serv;
using UcbBack.Models.Not_Mapped.ViewMoldes;
using UcbBack.Models.Serv;
using System.Data.Entity;
using System.Data.Entity.Migrations;
using UcbBack.Logic.B1;
using UcbBack.Models.Auth;

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

        [HttpGet]
        [Route("api/ServContract/History/")]
        public IHttpActionResult History()
        {
            var user = auth.getUser(Request);
            var query = "select * from " + CustomSchema.Schema + ".\"Serv_Process\" " +
                        " where \"State\" = '" + ServProcess.Serv_FileState.PendingApproval + "' " +
                        " or \"State\" = '" + ServProcess.Serv_FileState.INSAP + "' " +
                        " or \"State\" = '" + ServProcess.Serv_FileState.Rejected + "' " +
                        " order by (" +
                        "   case when \"State\" = '" + ServProcess.Serv_FileState.PendingApproval + "' then 1 " +
                        " when \"State\" = '" + ServProcess.Serv_FileState.INSAP + "' then 3 " +
                        " when \"State\" = '" + ServProcess.Serv_FileState.Rejected + "' then 5 " +
                        " end) asc, " +
                        " \"CreatedAt\" desc;";
            var rawresult = _context.Database.SqlQuery<ServProcess>(query).ToList();

            if (rawresult.Count() == 0)
                return NotFound();

            var res = auth.filerByRegional(rawresult.AsQueryable(), user).Cast<ServProcess>();

            if (res.Count() == 0)
                return Unauthorized();

            var res2 = (from r in res
                join b in _context.Branch.ToList()
                    on r.BranchesId equals b.Id
                select new
                {
                    r.Id,
                    r.BranchesId,
                    Branches = b.Name,
                    r.FileType,
                    r.State,
                    r.SAPId,
                    CreatedAt = r.CreatedAt.ToString("dd MMMM yyyy HH:mm")
                }).ToList();
            return Ok(res2);
        }

        [HttpGet]
        [Route("api/ServContract/{id}")]
        public IHttpActionResult Get(int id)
        {
            var user = auth.getUser(Request);
            var rawresult = _context.ServProcesses.Where(x=>x.Id==id);

            if (rawresult.Count() == 0)
                return NotFound();

            var res = auth.filerByRegional(rawresult, user).Cast<ServProcess>();

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
        public IHttpActionResult DeleteFile(JObject data)
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

                var user = auth.getUser(Request);

                int userid = Int32.Parse(Request.Headers.GetValues("id").First());
                ServProcess file = AddFileToProcess(Int32.Parse(o.BranchesId.ToString()), o.FileType.ToString(), userid);

                if (file == null)
                {
                    response.StatusCode = HttpStatusCode.BadRequest;
                    response.Headers.Add("UploadErrors", "{ \"Ya se Subio archivos para este mes\": \"Ya se subio  datos para este mes, si quiere volver a subir cancele el anterior archivo.\"}");
                    response.Content = new StringContent("Ya se subió  datos para este mes, si quiere volver a subir cancele el anterior archivo.");
                    return response;
                }

                DynamicExcelToDB(o.FileType,o,file,user,out response);
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
                    response.Headers.Add("UploadErrors", "{ \"La conexion con SAP se perdio\": \"No se pudo validar el archivo con con SAP\"}");
                    response.Content = new StringContent("Error conexion SAP");
                    return response;
                }
                Console.WriteLine(e);
                response.StatusCode = HttpStatusCode.BadRequest;
                response.Headers.Add("UploadErrors", "{ \"La conexion con SAP se perdio\": \"No se pudo validar el archivo con con SAP\"}");
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
            string query = null;
            switch (process.FileType)
            {
                case ServProcess.Serv_FileType.Varios:
                    query =
                        "select sv.\"CardName\", ou.\"Cod\" as \"OU\", sv.\"PEI\", sv.\"ServiceName\" as \"Memo\",  " +
                        " sv.\"ContractObjective\" as \"LineMemo\", sv.\"AssignedAccount\", sv.\"TotalAmount\" as \"Debit\"" +
                        " from \"ADMNALRRHHOLD\".\"Serv_Varios\" sv " +
                        " inner join \"ADMNALRRHHOLD\".\"Dependency\" d " +
                        " on sv.\"DependencyId\" = d.\"Id\" " +
                        " inner join \"ADMNALRRHHOLD\".\"OrganizationalUnit\" ou " +
                        " on d.\"OrganizationalUnitId\" = ou.\"Id\" " +
                        " where \"Serv_ProcessId\" = " + process.Id +
                        " order by sv.\"Id\" asc;";

                    break;
                case ServProcess.Serv_FileType.Carrera:
                    query =
                        "select sv.\"CardName\", ou.\"Cod\" as \"OU\", sv.\"PEI\", sv.\"ServiceName\" as \"Memo\",  " +
                        " sv.\"AssignedJob\"||\' \'||sv.\"Carrera\"||\' \'||sv.\"Student\" as \"LineMemo\", sv.\"AssignedAccount\", sv.\"TotalAmount\" as \"Debit\"" +
                        " from \"ADMNALRRHHOLD\".\"Serv_Carrera\" sv " +
                        " inner join \"ADMNALRRHHOLD\".\"Dependency\" d " +
                        " on sv.\"DependencyId\" = d.\"Id\" " +
                        " inner join \"ADMNALRRHHOLD\".\"OrganizationalUnit\" ou " +
                        " on d.\"OrganizationalUnitId\" = ou.\"Id\"" +
                        " where \"Serv_ProcessId\" = " + process.Id +
                        " order by sv.\"Id\" asc;";

                    break;
                case ServProcess.Serv_FileType.Proyectos:
                    query =
                        "select sv.\"CardName\", ou.\"Cod\" as \"OU\", sv.\"PEI\", sv.\"ServiceName\" as \"Memo\",  " +
                        " sv.\"ProjectSAPName\" as \"LineMemo\", sv.\"AssignedAccount\", sv.\"TotalAmount\" as \"Debit\"" +
                        " from \"ADMNALRRHHOLD\".\"Serv_Proyectos\" sv " +
                        " inner join \"ADMNALRRHHOLD\".\"Dependency\" d " +
                        " on sv.\"DependencyId\" = d.\"Id\" " +
                        " inner join \"ADMNALRRHHOLD\".\"OrganizationalUnit\" ou " +
                        " on d.\"OrganizationalUnitId\" = ou.\"Id\"" +
                        " where \"Serv_ProcessId\" = " + process.Id +
                        " order by sv.\"Id\" asc;";

                    break;
                case ServProcess.Serv_FileType.Paralelo:
                    query =
                        "select sv.\"CardName\", ou.\"Cod\" as \"OU\", sv.\"PEI\", sv.\"ServiceName\" as \"Memo\",  " +
                        " sv.\"Sigla\" as \"LineMemo\", sv.\"AssignedAccount\", sv.\"TotalAmount\" as \"Debit\"" +
                        " from \"ADMNALRRHHOLD\".\"Serv_Paralelo\" sv " +
                        " inner join \"ADMNALRRHHOLD\".\"Dependency\" d " +
                        " on sv.\"DependencyId\" = d.\"Id\" " +
                        " inner join \"ADMNALRRHHOLD\".\"OrganizationalUnit\" ou " +
                        " on d.\"OrganizationalUnitId\" = ou.\"Id\"" +
                        " where \"Serv_ProcessId\" = " + process.Id +
                        " order by sv.\"Id\" asc;";

                    break;
            }
            if (query == null)
                return NotFound();

            IEnumerable<Serv_Voucher> voucher = _context.Database.SqlQuery<Serv_Voucher>(query).ToList();

            return Ok(voucher);
        }

        [HttpPost]
        [Route("api/ServContract/CheckUpload")]
        public IHttpActionResult CheckUpload([FromBody] JObject upload)
        {
            int branchid = 0;
            int processid = 0;
            if (upload["FileType"] == null || upload["BranchesId"] == null || !Int32.TryParse(upload["BranchesId"].ToString(), out branchid) || upload["ProcessId"] == null)
                return BadRequest("Debes enviar Tipo de Archivo y segmentoOrigen");



            var FileType = upload["FileType"].ToString();

            ServProcess process = null;


            if (Int32.TryParse(upload["ProcessId"].ToString(), out processid))
            {
                process = _context.ServProcesses.FirstOrDefault(f => f.BranchesId == branchid
                                                                     && f.Id == processid);
            }
            else
            {
                process = _context.ServProcesses.FirstOrDefault(f => f.BranchesId == branchid
                                                                     && (f.State == ServProcess.Serv_FileState.Started)
                                                                     && f.FileType == FileType);
            }
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


        [HttpGet]
        [Route("api/ServContract/GetDistribution/{id}")]
        public HttpResponseMessage GetDistribution(int id)
        {
            HttpResponseMessage response = new HttpResponseMessage();

            var process = _context.ServProcesses.Include(x=>x.Branches).FirstOrDefault(p => p.Id == id);

            if (process == null)
            {
                response.StatusCode = HttpStatusCode.NotFound;
                return response;
            }
            var ex = new XLWorkbook();
            var d = new Distribution();
            switch (process.FileType)
            {
                case ServProcess.Serv_FileType.Varios:
                    var dist = _context.ServVarioses.Include(x=>x.Dependency).Include(x=>x.Dependency.OrganizationalUnit).
                        Where(x => x.Serv_ProcessId == process.Id).Select(x=>new
                    {
                        Id = x.Id,
                        Codigo_Socio=x.CardCode,
                        Nombre_Socio=x.CardName,
                        Cod_Dependencia=x.Dependency.Cod,
                        Cod_UO = x.Dependency.OrganizationalUnit.Cod,
                        PEI_PO=x.PEI,
                        Nombre_del_Servicio=x.ServiceName,
                        Objeto_del_Contrato=x.ContractObjective,
                        Cuenta_Asignada=x.AssignedAccount,
                        Monto_Contrato=x.ContractAmount,
                        Monto_IUE=x.IUE,
                        Monto_IT=x.IT,
                        Monto_a_Pagar=x.TotalAmount,
                        Observaciones=x.Comments,
                    }).OrderBy(x=>x.Id);
                    ex.Worksheets.Add(d.CreateDataTable(dist), "TotalDetalle");
                    break;
                case ServProcess.Serv_FileType.Carrera:
                    var dist1 = _context.ServCarreras.Include(x => x.Dependency).Include(x => x.Dependency.OrganizationalUnit).
                        Where(x => x.Serv_ProcessId == process.Id).Select(x=>new
                    {
                        Id = x.Id,
                        Codigo_Socio = x.CardCode,
                        Nombre_Socio = x.CardName,
                        Cod_Dependencia = x.Dependency.Cod,
                        Cod_UO = x.Dependency.OrganizationalUnit.Cod,
                        PEI_PO = x.PEI,
                        Nombre_del_Servicio = x.ServiceName,
                        Codigo_Carrera=x.Carrera,
                        Documento_Base=x.DocumentNumber,
                        Postulante=x.Student,
                        Tipo_Tarea_Asignada=x.AssignedJob,
                        Cuenta_Asignada = x.AssignedAccount,
                        Monto_Contrato = x.ContractAmount,
                        Monto_IUE = x.IUE,
                        Monto_IT = x.IT,
                        Monto_a_Pagar = x.TotalAmount,
                        Observaciones = x.Comments,
                    }).OrderBy(x => x.Id);
                    ex.Worksheets.Add(d.CreateDataTable(dist1), "TotalDetalle");
                    break;
                case ServProcess.Serv_FileType.Paralelo:
                    var dist2 = _context.ServParalelos.Include(x => x.Dependency).Include(x => x.Dependency.OrganizationalUnit).
                        Where(x => x.Serv_ProcessId == process.Id).Select(x=>new
                    {
                        Id = x.Id,
                        Codigo_Socio = x.CardCode,
                        Nombre_Socio = x.CardName,
                        Cod_Dependencia = x.Dependency.Cod,
                        Cod_UO = x.Dependency.OrganizationalUnit.Cod,
                        PEI_PO = x.PEI,
                        Nombre_del_Servicio = x.ServiceName,
                        Periodo_Academico = x.Periodo,
                        Sigla_Asignatura = x.Sigla,
                        Paralelo = x.ParalelNumber,
                        Codigo_Paralelo_SAP = x.ParalelSAP,
                        Cuenta_Asignada = x.AssignedAccount,
                        Monto_Contrato = x.ContractAmount,
                        Monto_IUE = x.IUE,
                        Monto_IT = x.IT,
                        Monto_a_Pagar = x.TotalAmount,
                        Observaciones = x.Comments,
                    }).OrderBy(x => x.Id);

                    ex.Worksheets.Add(d.CreateDataTable(dist2), "TotalDetalle");
                    break;
                case ServProcess.Serv_FileType.Proyectos:
                    var dist3 = _context.ServProyectoses.Include(x => x.Dependency).Include(x => x.Dependency.OrganizationalUnit).
                        Where(x => x.Serv_ProcessId == process.Id).Select(x => new
                    {
                        Id = x.Id,
                        Codigo_Socio = x.CardCode,
                        Nombre_Socio = x.CardName,
                        Cod_Dependencia = x.Dependency.Cod,
                        Cod_UO = x.Dependency.OrganizationalUnit.Cod,
                        PEI_PO = x.PEI,
                        Nombre_del_Servicio = x.ServiceName,
                        Codigo_Proyecto_SAP=x.ProjectSAPCode,
                        Nombre_del_Proyecto=x.ProjectSAPName,
                        x.Version,
                        Periodo_Academico = x.Periodo,
                        Tipo_Tarea_Asignada = x.AssignedJob,
                        Cuenta_Asignada = x.AssignedAccount,
                        Monto_Contrato = x.ContractAmount,
                        Monto_IUE = x.IUE,
                        Monto_IT = x.IT,
                        Monto_a_Pagar = x.TotalAmount,
                        Observaciones = x.Comments,
                    }).OrderBy(x => x.Id);
                    ex.Worksheets.Add(d.CreateDataTable(dist3), "TotalDetalle");
                    break;
            }
            var ms = new MemoryStream();
            ex.SaveAs(ms);
            response.StatusCode = HttpStatusCode.OK;
            response.Content = new StreamContent(ms);
            response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");
            response.Content.Headers.ContentDisposition.FileName = process.Branches.Abr + "-Lote_" + process.Id + "-" + process.FileType + ".xlsx";
            response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            response.Content.Headers.ContentLength = ms.Length;
            ms.Seek(0, SeekOrigin.Begin);
            return response;
        }
        [HttpGet]
        [Route("api/ServContractprocessRows/{id}")]
        public IHttpActionResult GetSAPResumeRows(int id)
        {
            var processes = _context.ServProcesses.Include(x => x.Branches).FirstOrDefault(f =>
                f.Id == id);
            if (processes == null)
            {
                return NotFound();
            }
            var data = processes.getVoucherData(_context);
            var ppagar = data.Where(g => g.Concept == "PPAGAR").Select(g => new Serv_Voucher()
            {
                CardName = g.CardName,
                CardCode = g.CardCode,
                OU = g.OU,
                PEI = g.PEI,
                Carrera = g.Carrera,
                Paralelo = g.Paralelo,
                Periodo = g.Periodo,
                ProjectCode = g.ProjectCode,
                Memo = g.Memo,
                LineMemo = g.LineMemo,
                Concept = g.Concept,
                AssignedAccount = g.AssignedAccount,
                Account = g.Account,
                Credit = g.Credit,
                Debit = g.Debit
            }).ToList();

            List<Serv_Voucher> rest = data.Where(g => g.Concept != "PPAGAR").GroupBy(g => new
            {
                g.CardCode,
                g.OU,
                g.PEI,
                g.Carrera,
                g.Paralelo,
                g.Periodo,
                g.ProjectCode,
                g.Memo,
                g.LineMemo,
                g.Concept,
                g.AssignedAccount,
                g.Account,
            }).Select(g => new Serv_Voucher()
            {
                CardName = "",
                CardCode = g.Key.CardCode,
                OU = g.Key.OU,
                PEI = g.Key.PEI,
                Carrera = g.Key.Carrera,
                Paralelo = g.Key.Paralelo,
                Periodo = g.Key.Periodo,
                ProjectCode = g.Key.ProjectCode,
                Memo = g.Key.Memo,
                LineMemo = g.Key.LineMemo,
                Concept = g.Key.Concept,
                AssignedAccount = g.Key.AssignedAccount,
                Account = g.Key.Account,
                Credit = g.Sum(s => s.Credit),
                Debit = g.Sum(s => s.Debit)
            }).ToList();

            List<Serv_Voucher> dist1 = ppagar.Union(rest).OrderBy(z => z.Debit == 0.00M ? 1 : 0).ThenBy(z => z.Account).ToList();

            dynamic res = new JObject();

            res.rowCount = dist1.Count();
            return Ok(res);
        }

        [HttpGet]
        [Route("api/ServContractToApproval/{id}")]
        public IHttpActionResult ToApproval(int id)
        {
            var user = auth.getUser(Request);
            var processes = _context.ServProcesses.Where(f =>
                f.Id == id && f.State == ServProcess.Serv_FileState.Started);
            if (processes.Count() == 0)
            {
                return NotFound();
            }

            processes = auth.filerByRegional(processes, user).Cast<ServProcess>();
            var process = processes.FirstOrDefault();

            if (process==null)
                return Unauthorized();

            process.State = ServProcess.Serv_FileState.PendingApproval;
            _context.ServProcesses.AddOrUpdate(process);
            _context.SaveChanges();

            return Ok();
        }

        [HttpDelete]
        [Route("api/ServContract/{id}")]
        public IHttpActionResult DeleteProcess(int id)
        {
            var user = auth.getUser(Request);
            var processes = _context.ServProcesses.Where(x =>
                x.Id == id && (x.State == ServProcess.Serv_FileState.Started || x.State == ServProcess.Serv_FileState.PendingApproval));
            if (processes.Count() == 0)
                return NotFound();

            processes = auth.filerByRegional(processes, user).Cast<ServProcess>();
            var process = processes.FirstOrDefault();

            if (process == null)
                return Unauthorized();

            switch (process.State)
            {
                case ServProcess.Serv_FileState.Started:
                    process.State = ServProcess.Serv_FileState.Canceled;
                    break;
                case ServProcess.Serv_FileState.PendingApproval:
                    process.State = ServProcess.Serv_FileState.Rejected;
                    break;
            }
            process.LastUpdatedBy = user.Id;
            _context.ServProcesses.AddOrUpdate(process);
            _context.SaveChanges();
            return Ok();
        }

        [HttpPost]
        [Route("api/ServContractToSAP/{id}")]
        public IHttpActionResult ToSAP(int id, JObject webdata)
        {
            if (webdata == null || webdata["date"] == null)
            {
                return BadRequest();
            }

            var B1 = B1Connection.Instance();
            HttpResponseMessage response = new HttpResponseMessage();
            var user = auth.getUser(Request);
            var processes = _context.ServProcesses.Include(x=>x.Branches).Where(f =>
                f.Id == id && f.State == ServProcess.Serv_FileState.PendingApproval);

            if (processes.Count() == 0)
            {
                return NotFound();
            }

            processes = auth.filerByRegional(processes, user).Cast<ServProcess>();
            var process = processes.FirstOrDefault();

            if (process == null)
            {
                return Unauthorized();
            }

            DateTime date = DateTime.Parse(webdata["date"].ToString());
            process.InSAPAt = date;
            var data = process.getVoucherData(_context);
            var memos = data.Select(x => x.Memo).Distinct().ToList();

            foreach (var memo in memos)
            {
                //remove special chars
                var goodMemo = Regex.Replace(memo, "[^\\w\\._]", "");
                //remove new line characters
                goodMemo = Regex.Replace(goodMemo, @"\t|\n|\r", "");

                var ppagar = data.Where(g => g.Concept == "PPAGAR" && g.Memo == memo).Select(g => new Serv_Voucher()
                {
                    CardName=g.CardName,
                    CardCode=g.CardCode,
                    OU=g.OU,
                    PEI=g.PEI,
                    Carrera=g.Carrera,
                    Paralelo=g.Paralelo,
                    Periodo=g.Periodo,
                    ProjectCode=g.ProjectCode,
                    Memo=g.Memo,
                    LineMemo=g.LineMemo,
                    Concept=g.Concept,
                    AssignedAccount=g.AssignedAccount,
                    Account=g.Account,
                    Credit = g.Credit,
                    Debit = g.Debit
                }).ToList();

                List<Serv_Voucher> rest = data.Where(g => g.Concept != "PPAGAR" && g.Memo == memo).GroupBy(g => new
                {
                    g.CardCode,
                    g.OU,
                    g.PEI,
                    g.Carrera,
                    g.Paralelo,
                    g.Periodo,
                    g.ProjectCode,
                    g.Memo,
                    g.LineMemo,
                    g.Concept,
                    g.AssignedAccount,
                    g.Account,
                }).Select(g => new Serv_Voucher()
                {
                    CardName = "",
                    CardCode=g.Key.CardCode,
                    OU=g.Key.OU,
                    PEI=g.Key.PEI,
                    Carrera=g.Key.Carrera,
                    Paralelo=g.Key.Paralelo,
                    Periodo=g.Key.Periodo,
                    ProjectCode=g.Key.ProjectCode,
                    Memo=g.Key.Memo,
                    LineMemo=g.Key.LineMemo,
                    Concept=g.Key.Concept,
                    AssignedAccount=g.Key.AssignedAccount,
                    Account=g.Key.Account,
                    Credit = g.Sum(s => s.Credit),
                    Debit = g.Sum(s => s.Debit)
                }).ToList();

                List<Serv_Voucher> dist1 = ppagar.Union(rest).OrderBy(z => z.Debit == 0.00M ? 1 : 0).ThenBy(z => z.Account).ToList();
                B1.addServVoucher(user.Id,dist1.ToList(),process);
            }

            if(memos.Count()>1)
                process.SAPId = "Multiples.";
            process.State = ServProcess.Serv_FileState.INSAP;
            process.LastUpdatedBy = user.Id;
            _context.ServProcesses.AddOrUpdate(process);
            _context.SaveChanges();

            return Ok(process.SAPId);
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
            realfileName = Abr + "-CC_" + fileType;
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
        private void DynamicExcelToDB(string FileType, dynamic o, ServProcess file,CustomUser user,  out HttpResponseMessage response)
        {
            response = new HttpResponseMessage();
            switch (FileType)
            {
                case ServProcess.Serv_FileType.Varios:
                    Serv_VariosExcel ExcelFile = new Serv_VariosExcel(o.excelStream, _context, o.fileName,file,user,headerin:1,sheets:1);
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
                    Serv_CarreraExcel ExcelFile2 = new Serv_CarreraExcel(o.excelStream, _context, o.fileName, file,user, headerin: 1, sheets: 1);
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
                    Serv_ProyectosExcel ExcelFile3 = new Serv_ProyectosExcel(o.excelStream, _context, o.fileName, file, user,headerin: 1, sheets: 1);
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
                    Serv_ParaleloExcel ExcelFile4 = new Serv_ParaleloExcel(o.excelStream, _context, o.fileName, file,user, headerin: 1, sheets: 1);
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
