using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Data.Entity.Migrations.History;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Http;
using UcbBack.Models;
using UcbBack.Logic;
using System.Globalization;
using Microsoft.Ajax.Utilities;
using Newtonsoft.Json.Linq;
using UcbBack.Logic.B1;
using UcbBack.Models.Not_Mapped.CustomDataAnnotations;
using UcbBack.Models.Not_Mapped.ViewMoldes;


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
                        " where (\"Active\"=true or \"EndDate\">=current_date) ;";
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
                    Linkagestr = x.Link.Value,
                    Linkage = x.Linkage,
                    StartDatestr = x.StartDate.ToString("dd MMM yyyy", new CultureInfo("es-ES")),
                    EndDatestr = x.EndDate == null ? "" : x.EndDate.GetValueOrDefault().ToString("dd MMM yyyy", new CultureInfo("es-ES")),
                    StartDate = x.StartDate.ToString("MM/dd/yyyy"),
                    EndDate = x.EndDate == null ? "" : x.EndDate.Value.ToString("MM/dd/yyyy"),
                    x.NumGestion,
                    x.Comunicado,
                    x.Respaldo,
                    x.Seguimiento
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

        // todo convert this to a report in excel
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
                contract.PositionDescription = contract.PositionDescription!=null?contract.PositionDescription.ToUpper():null;
                contract.Active = true;
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
            contractInDB.Active = false;
            contractInDB.UpdatedAt = DateTime.Now;
            _context.SaveChanges();
            return Ok(contractInDB);
        }
        //Bajas
        // GET api/Contract/BajaPendiente/
        [HttpGet]
        [Route("api/Contract/BajaPendiente")]
        public IHttpActionResult BajaPendiente()
        {
            var query = "select * from " + CustomSchema.Schema + ".lastcontracts " +
                        " where \"EndDate\" is not null and year(\"EndDate\")*100+month(\"EndDate\") <= year(current_date)*100+month(current_date) and \"Active\" = true ";
            var rawresult = _context.Database.SqlQuery<ContractDetailViewModel>(query).ToList();

            var user = auth.getUser(Request);

            var res = auth.filerByRegional(rawresult.AsQueryable(), user);

            return Ok(res);
        }

        // GET api/Contract/BajaPendiente/
        [HttpPost]
        [Route("api/Contract/ConfirmBajaPendiente")]
        public IHttpActionResult ConfirmBajaPendiente(JObject data)
        {
            var list = data.Values().ToList()[0].ToList();
            foreach (var d in list)
            {
                var id = Int32.Parse(d.ToString());
                var contract = _context.ContractDetails.FirstOrDefault(x => x.Id == id);
                if (contract != null)
                {
                    contract.Cause = "5";
                    contract.Active = false;
                    contract.UpdatedAt = DateTime.Now;
                }
            }
            _context.SaveChanges();
            return Ok();
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
            contractInDB.PositionDescription = contract.PositionDescription == null ? null : contract.PositionDescription.ToUpper();
            contractInDB.Linkage = contract.Linkage;
            contractInDB.AI = contract.AI;
            contractInDB.NumGestion = contract.NumGestion;
            contractInDB.Seguimiento = contract.Seguimiento;
            contractInDB.Respaldo = contract.Respaldo;
            contractInDB.Comunicado = contract.Comunicado;
            contractInDB.UpdatedAt = DateTime.Now;

            var person = _context.Person.FirstOrDefault(x => x.CUNI == contractInDB.CUNI);

            var user = auth.getUser(Request);
            // create user in SAP
            B1.AddOrUpdatePerson(user.Id, person);

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
