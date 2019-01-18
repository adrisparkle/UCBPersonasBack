using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Data.Entity.Migrations.History;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web;
using System.Web.Http;
using UcbBack.Models;
using ExcelDataReader;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using UcbBack.Logic;
using UcbBack.Logic.ExcelFiles;
using System.Globalization;
using Microsoft.Ajax.Utilities;
using Newtonsoft.Json.Linq;
using UcbBack.Logic.B1;
using UcbBack.Models.Not_Mapped.CustomDataAnnotations;
using UcbBack.Models.ViewMoldes;


namespace UcbBack.Controllers
{
    public class ContractController : ApiController
    {
         private ApplicationDbContext _context;
        private ValidateAuth auth;
        private B1Connection B1;
        private ADClass AD;

        public ContractController()
        {

            _context = new ApplicationDbContext();
            auth = new ValidateAuth();
            B1 = B1Connection.Instance();
            AD = new ADClass();

        }

        // GET api/Contract
        [Route("api/Contract")]
        public IHttpActionResult Get()
        {
            var query ="select * from " + CustomSchema.Schema + ".lastcontracts "+
                        " where \"EndDate\" is null or \"EndDate\"> current_date";
            var rawresult = _context.Database.SqlQuery<ContractDetailViewModel>(query).ToList();

            var user = auth.getUser(Request);

            var res = auth.filerByRegional(rawresult.AsQueryable(), user);

            return Ok(res);
        }

        [HttpGet]
        [Route("api/ContractSAP")]
        public IHttpActionResult GetSAP()
        {
            DateTime date = DateTime.Now;
            var contplist = _context.ContractDetails
                .Include(p => p.Branches)
                .Include(p => p.Dependency)
                .Include(p => p.Positions)
                .Include(p => p.People).ToList()
                .Where(x => /*x.StartDate <= date
                            && */(x.EndDate == null || x.EndDate.Value.Year * 100 + x.EndDate.Value.Month >= date.Year * 100 + date.Month))
                .OrderByDescending(x => x.StartDate)
                .ToList()
                .Select(x => new
                {
                    x.People.CUNI,
                    x.People.Document,
                    FullName = x.People.GetFullName(),
                    Dependency = x.Dependency.Name,
                    DependencyCod = x.Dependency.Cod,
                    x.BranchesId
                }).ToList();
            var user = auth.getUser(Request);

            var res = auth.filerByRegional(contplist.AsQueryable(), user);

            return Ok(res);
        }

        [HttpGet]
        [Route("api/Contract/GetPersonHistory/{id}")]
        public IHttpActionResult GetHistory(int id,[FromUri] string all = "true")
        {
            if (all=="true")
            {
                var history = from contract in _context.ContractDetails.Include(x => x.Dependency)
                        .Include(x => x.Positions)
                        .Include(x => x.Link)
                        .Where(x => x.PeopleId == id).ToList()
                    join brs in _context.Branch on contract.Dependency.BranchesId equals brs.Id
                    select new
                    {
                        contract.Id,
                        contract.DependencyId,
                        contract.Dependency.Cod,
                        Dependency = contract.Dependency.Name,
                        contract.Dependency.BranchesId,
                        Branches = brs.Name,
                        Positions = contract.Positions.Name,
                        contract.PositionsId,
                        contract.PositionDescription,
                        contract.Dedication,
                        Linkage = contract.Link.Value,
                        contract.StartDate,
                        contract.EndDate
                    };
                return Ok(history);
            }
            else
            {
                DateTime date = DateTime.Now;
                var history = from contract in _context.ContractDetails.Include(x => x.Dependency)
                        .Include(x => x.Positions)
                        .Include(x => x.Link)
                        .Where(x => x.PeopleId == id).ToList()
                    join brs in _context.Branch on contract.Dependency.BranchesId equals brs.Id
                              where (contract.EndDate == null || contract.EndDate.Value.Year * 100 + contract.EndDate.Value.Month >= date.Year * 100 + date.Month)
                    select new
                    {
                        contract.Id,
                        contract.DependencyId,
                        contract.Dependency.Cod,
                        Dependency = contract.Dependency.Name,
                        contract.Dependency.BranchesId,
                        Branches = brs.Name,
                        Positions = contract.Positions.Name,
                        contract.PositionsId,
                        contract.PositionDescription,
                        contract.Dedication,
                        Linkage = contract.Link.Value,
                        contract.StartDate,
                        contract.EndDate
                    };
                return Ok(history);
            }   
        }

        // GET api/Contract
        [Route("api/Contract/{id}")]
        public IHttpActionResult GetContract(int id)
        {
            DateTime date = DateTime.Now;
            var contplist = _context.ContractDetails
                .Include(p => p.Branches)
                .Include(p => p.Dependency)
                .Include(p => p.Positions)
                .Include(p => p.People)
                .Include(p => p.Link)
                .Where(x => x.Id == id)
                .OrderByDescending(x => x.StartDate)
                .ToList()
                .Select(x => new
                {
                    x.Id,
                    x.People.CUNI,
                    PeopleId = x.People.Id,
                    x.People.Document,
                    FullName = x.People.GetFullName(),
                    DependencyId = x.Dependency.Id,
                    Dependency = x.Dependency.Name,
                    Branches = x.Branches.Abr,
                    BranchesId = x.Branches.Id,
                    PositionsId = x.Positions.Id,
                    Positions = x.Positions.Name,
                    x.PositionDescription,
                    x.AI,
                    x.Dedication,
                    Linkage = x.Link.Value,
                    //StartDate = x.StartDate.ToString("dd MMM yyyy", new CultureInfo("es-ES")),
                    //EndDate = x.EndDate == null ? "" : x.EndDate.GetValueOrDefault().ToString("dd MMM yyyy", new CultureInfo("es-ES"))
                    StartDate = x.StartDate.ToString("MM/dd/yyyy"),
                    EndDate = x.EndDate == null ? "" : x.EndDate.Value.ToString("MM/dd/yyyy")
                });

            var user = auth.getUser(Request);
            var res = auth.filerByRegional(contplist.AsQueryable(), user);
            if (res.Count() == 0)
                return NotFound();

            return Ok(res.FirstOrDefault());
        }

        [HttpGet]
        [Route("api/Contract/GetPersonContract/{id}")]
        public IHttpActionResult GetPersonContract(int id)
        {
            List<ContractDetail> contractInDB = null;

            contractInDB = _context.ContractDetails.Where(d => d.People.Id == id).ToList();

            if (contractInDB == null)
                return NotFound();

            return Ok(contractInDB);
        }
        [HttpGet]
        [Route("api/Contract/GetContractsBranch/{id}")]
        public IHttpActionResult GetContractsBranch(int id)
        {
            List<ContractDetail> contractInDB = null;
            DateTime date=new DateTime(2018,9,1);
            DateTime date2=new DateTime(2019,1,1);
            var people = _context.ContractDetails.Include(x=>x.People).Include(x=>x.Branches).Include(x=>x.Link).Where(x=>  (x.EndDate==null || x.EndDate>date2)).Select(x=>x.People).Distinct();
            // var people = _context.CustomUsers.Include(x => x.People).Select(x => x.People);
            int i = people.Count();
            string res = "";
            var br = _context.Branch.ToList();
            var OU = _context.OrganizationalUnits.ToList();

            foreach (var person in people)
            {
                var contract = person.GetLastContract();
                var user = _context.CustomUsers.FirstOrDefault(x => x.PeopleId == contract.People.Id);
               /* res += contract.People.GetFullName() + ";";
                res += user.UserPrincipalName + ";";
                res += "NORMAL;";
                res += "NO;";
                res += contract.Branches.Abr + ";";
                res += "RENDICIONES;";
                res += contract.CUNI + ";";*/

                res += contract.People.CUNI + ";";
                res += contract.People.Document + ";";
                res += contract.People.GetFullName() + ";";
                res += contract.People.FirstSurName + ";";
                res += contract.People.SecondSurName + ";";
                res += contract.People.MariedSurName + ";";
                res += contract.People.Names + ";";
                res += contract.People.BirthDate + ";";
                res += contract.People.UcbEmail + ";";


                res += contract.Dependency.Cod + ";";
                res += contract.Dependency.Name + ";";

                var o = OU.FirstOrDefault(x => x.Id == contract.Dependency.OrganizationalUnitId);
                res += o.Cod + ";";
                res += o.Name + ";";

                res += contract.Positions.Name + ";";
                res += contract.Dedication + ";";
                res += contract.Link.Value + ";";
                res += contract.AI + ";";
                var lm = contract.People.GetLastManagerAuthorizator(_context);
                res += lm==null?";":lm.CUNI + ";";
                res += lm==null?";":lm.GetFullName() + ";";
                var lmlc = lm==null?null:lm.GetLastContract(_context);
                res += lmlc == null ? ";" : lmlc.Positions.Name + ";";
                res += lmlc == null ? ";" : lmlc.Dependency.Cod + ";";
                res += lmlc == null ? ";" : lmlc.Dependency.Name + ";";

                var b = br.FirstOrDefault(x => x.Id == contract.Dependency.BranchesId);
                res += b.Abr + ";";
                res += b.Name;

                res += "\n";
            }


            return Ok(res);
        }

        [NonAction]
        public async Task<System.Dynamic.ExpandoObject> HttpContentToVariables(MultipartMemoryStreamProvider req)
        {
            dynamic res = new System.Dynamic.ExpandoObject();
            foreach (HttpContent contentPart in req.Contents)
            {
                var contentDisposition = contentPart.Headers.ContentDisposition;
                string varname = contentDisposition.Name;
                if (varname == "\"segmentoOrigen\"")
                {
                    res.segmentoOrigen = contentPart.ReadAsStringAsync().Result;
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
        [Route("api/Contract/AltaExcel/save/{id}")]
        public IHttpActionResult saveLastAltaExcel(int id)
        {
            var tempAlta = _context.TempAltas.Where(x => x.BranchesId == id && x.State != "SAVED");
            if (tempAlta.Count() > 0)
                return NotFound();

            var validator = new ValidatePerson();

            foreach (var alta in tempAlta)
            {
                var person = new People();

                if (alta.State == "NEW")
                {
                    person.Id = People.GetNextId(_context);
                    person.FirstSurName = alta.FirstSurName.Trim();
                    person.SecondSurName = alta.SecondSurName.Trim().IsNullOrWhiteSpace() ? null : alta.SecondSurName.Trim();
                    person.MariedSurName = alta.MariedSurName.Trim().IsNullOrWhiteSpace() ? null : alta.MariedSurName.Trim();
                    person.Names = alta.Names.Trim();
                    person.BirthDate = alta.BirthDate;
                    person.Gender = alta.Gender;

                    person.AFP = alta.AFP;
                    person.NUA = alta.NUA;

                    person.Document = alta.Document;
                    person.Ext = alta.Ext;
                    person.TypeDocument = alta.TypeDocument;

                    person.UseSecondSurName = person.SecondSurName.IsNullOrWhiteSpace();
                    person.UseMariedSurName = person.MariedSurName.IsNullOrWhiteSpace();

                    person = validator.UcbCode(person);
                    person.Pending = true;

                    _context.Person.Add(person);
                }
                else
                {
                    person = _context.Person.FirstOrDefault(x => x.CUNI == alta.CUNI);
                }

                var contract = new ContractDetail();

                contract.Id = ContractDetail.GetNextId(_context);
                contract.DependencyId = _context.Dependencies.FirstOrDefault(x=>x.Cod==alta.Dependencia).Id;
                contract.CUNI = person.CUNI;
                contract.PeopleId = person.Id;
                contract.BranchesId = alta.BranchesId;
                contract.Dedication = "TH";
                contract.Linkage = 3;
                contract.PositionDescription = "Docente Tiempo Horario";
                contract.PositionsId = 26;
                contract.StartDate = alta.StartDate;
                contract.EndDate = alta.EndDate;

            }

            _context.SaveChanges();
            return Ok(tempAlta);
        }


        [HttpDelete]
        [Route("api/Contract/AltaExcel")]
        public IHttpActionResult removeLastAltaExcel(JObject data)
        {
            int branchesid;
            if (data["segmentoOrigen"] == null || !Int32.TryParse(data["segmentoOrigen"].ToString(), out branchesid))
            {
                ModelState.AddModelError("Mal Formato", "Debes enviar mes, gestion y segmentoOrigen");
                return BadRequest();

            }
            List<TempAlta> tempAlta = _context.TempAltas.Where(x => x.BranchesId == branchesid && x.State != "UPLOADED" && x.State != "CANCELED").ToList();
            foreach (var al in tempAlta)
            {
                al.State = "CANCELED";
            }

            _context.SaveChanges();
            return Ok();
        }

        [HttpGet]
        [Route("api/Contract/AltaExcel/{id}")]
        public IHttpActionResult getLastAltaExcel(int id)
        {
            List<TempAlta> tempAlta = _context.TempAltas.Where(x => x.BranchesId == id && x.State != "UPLOADED" && x.State != "CANCELED").ToList();
            return Ok(tempAlta);
        }

        [HttpGet]
        [Route("api/Contract/AltaExcel")]
        public HttpResponseMessage getAltaExcelTemplate()
        {
            ContractExcel contractExcel = new ContractExcel(fileName: "AltaExcel_TH.xlsx", headerin: 3);
            return contractExcel.getTemplate();
        }

        [HttpPost]
        [Route("api/Contract/AltaExcel")]
        public async Task<HttpResponseMessage> AltaExcel()
        {
            var response = new HttpResponseMessage();
            try
            {
                var req = await Request.Content.ReadAsMultipartAsync();
                dynamic o = HttpContentToVariables(req).Result;
                int segment = 0;
                if (o.segmentoOrigen == null || !Int32.TryParse(o.segmentoOrigen,out segment))
                {
                    response.StatusCode = HttpStatusCode.BadRequest;
                    response.Content = new StringContent("Debe enviar segmentoOrigen");
                    return response;
                }

                var segId = _context.Branch.FirstOrDefault(b => b.Id == segment);
                if (segId == null)
                {
                    response.StatusCode = HttpStatusCode.BadRequest;
                    response.Content = new StringContent("Debe enviar segmentoOrigen valido");
                    return response;
                }
                ContractExcel contractExcel = new ContractExcel(o.excelStream, _context, o.fileName, segId.Id, headerin: 3, sheets: 1);
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

        //altas
        // POST api/Contract/alta/4
        [HttpPost]
        [Route("api/Contract/Alta")]
        public IHttpActionResult Post([FromBody]ContractDetail contract)
        {
            if (!ModelState.IsValid)
                return BadRequest(ModelState);
            People person = _context.Person.FirstOrDefault(x => x.Id == contract.PeopleId);

            if (person == null)
                person = _context.Person.FirstOrDefault(x => x.CUNI == contract.CUNI);

            if (person == null)
                return NotFound();

            bool valid = true;
            string errorMessage = "";
            if (contract.EndDate < contract.StartDate)
            {
                valid = false;
                errorMessage += "La fecha fin no puede ser menor a la fecha inicio";
            }

            if (_context.Dependencies.FirstOrDefault(x=>x.Id == contract.DependencyId ) == null)
            {
                valid = false;
                errorMessage += "La Dependencia no es valida";
            }

            if (_context.Position.FirstOrDefault(x => x.Id == contract.PositionsId) == null)
            {
                valid = false;
                errorMessage += "El Cargo no es valido";
            }

            if ( _context.TableOfTableses.FirstOrDefault(x => x.Type == "VINCULACION" && x.Id == contract.Linkage) == null)
            {
                valid = false;
                errorMessage += "Esta vinculación no es valida";
            }

            if (_context.TableOfTableses.FirstOrDefault(x => x.Type == "DEDICACION" && x.Value == contract.Dedication) == null)
            {
                valid = false;
                errorMessage += "Esta dedicación no es valida";
            }

            if (valid)
            {
                contract.PeopleId = person.Id;
                contract.CUNI = person.CUNI;
                contract.BranchesId = _context.Dependencies.FirstOrDefault(x => x.Id == contract.DependencyId).BranchesId;
                contract.Id = ContractDetail.GetNextId(_context);
                _context.ContractDetails.Add(contract);
                _context.SaveChanges();

                var user = auth.getUser(Request);

                // create user in SAP
                B1.AddOrUpdatePerson(user.Id, person);

                return Created(new Uri(Request.RequestUri + "/" + contract.Id), contract);
            }

            return BadRequest(errorMessage);

        }

        //Bajas
        // POST api/Contract/Baja/5
        [HttpPost]
        [Route("api/Contract/Baja/{id}")]
        public IHttpActionResult Baja(int id, ContractDetail contract)
        {
            ContractDetail contractInDB = _context.ContractDetails.FirstOrDefault(d => d.Id == id);
            // contractInDB.EndDate=DateTime.Now;
            ChangesLogs log = new ChangesLogs();
            log.AddChangesLog(contractInDB, contract, new List<string>() { "EndDate", "Cause" });
            contractInDB.EndDate = contract.EndDate;
            contractInDB.Cause = contract.Cause;
            contractInDB.UpdatedAt = DateTime.Now;
            _context.SaveChanges();
            return Ok(contractInDB);
        }

        // PUT api/Contract/5
        [HttpPut]
        [Route("api/Contract/{id}")]
        public IHttpActionResult Put(int id, ContractDetail contract)
        {
            if (!ModelState.IsValid)
                return BadRequest(ModelState);

            ContractDetail contractInDB = _context.ContractDetails.FirstOrDefault(d => d.Id == id);
            if (contractInDB == null)
                return NotFound();
            // log changes
            ChangesLogs log = new ChangesLogs();
            log.AddChangesLog(contractInDB, contract, new List<string>() { "EnDate", "StartDate", "Dedication",  "DependencyId", "PositionsId", "PositionDescription", "Linkage", "AI" });
            // todo view rol and permisions to update or not
            contractInDB.StartDate = contract.StartDate;
            contractInDB.EndDate = contract.EndDate;
            contractInDB.Dedication = contract.Dedication;
            contractInDB.BranchesId = _context.Dependencies.FirstOrDefault(x=>x.Id==contract.DependencyId).BranchesId;
            contractInDB.DependencyId = contract.DependencyId;
            contractInDB.PositionsId = contract.PositionsId;
            contractInDB.PositionDescription = contract.PositionDescription;
            contractInDB.Linkage = contract.Linkage;
            contractInDB.AI = contract.AI;
            contractInDB.UpdatedAt = DateTime.Now;

            _context.SaveChanges();
            return Ok(contractInDB);
        }

        // DELETE api/Contract/5
        [HttpDelete]
        public IHttpActionResult Delete(int id)
        {
            var contractInDB = _context.Contracts.FirstOrDefault(d => d.Id == id);
            if (contractInDB == null)
                return NotFound();
            //_context.Contracts.Remove(contractInDB);
            //_context.SaveChanges();
            return Ok();
        }
    }
}
