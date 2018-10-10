using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
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


namespace UcbBack.Controllers
{
    public class ContractController : ApiController
    {
         private ApplicationDbContext _context;
        private ValidateAuth auth;

        public ContractController()
        {

            _context = new ApplicationDbContext();
            auth = new ValidateAuth();

        }

        // GET api/Contract
        [Route("api/Contract")]
        public IHttpActionResult Get()
        {
            DateTime date = DateTime.Now;
            var contplist = _context.ContractDetails
                .Include(p => p.Branches)
                .Include(p => p.Dependency)
                .Include(p => p.Positions)
                .Include(p => p.People)
                .Where(x => /*x.StartDate <= date
                            && */(x.EndDate == null || x.EndDate > date))
                .OrderByDescending(x=>x.StartDate)
                .ToList()
                .Select(x => new
                {
                    x.Id, 
                    x.People.CUNI, 
                    x.People.Document, 
                    FullName= x.People.GetFullName(),
                    Dependency = x.Dependency.Name, 
                    DependencyCod = x.Dependency.Cod, 
                    Branches = x.Branches.Abr, 
                    BranchesId = x.Branches.Id, 
                    Positions=x.Positions.Name, 
                    x.Dedication,
                    x.Linkage,
                    StartDate = x.StartDate.ToString("dd MMM yyyy", new CultureInfo("es-ES")),
                    EndDate = x.EndDate == null?null:x.EndDate.GetValueOrDefault().ToString("dd MMM yyyy", new CultureInfo("es-ES"))
                }).ToList();
            var user = auth.getUser(Request);

            var res = auth.filerByRegional(contplist.AsQueryable(), user);

            return Ok(res);
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
                .Where(x =>/* x.StartDate <= date
                            &&*/ (x.EndDate == null || x.EndDate > date)
                            && x.Id == id)
                .OrderByDescending(x => x.StartDate)
                .ToList()
                .Select(x => new
                {
                    x.Id,
                    x.People.CUNI,
                    x.People.Document,
                    FullName = x.People.GetFullName(),
                    Dependency = x.Dependency.Name,
                    Branches = x.Branches.Abr,
                    BranchesId = x.Branches.Id,
                    Positions = x.Positions.Name,
                    x.Dedication,
                    x.Linkage,
                    StartDate = x.StartDate.ToString("dd MMM yyyy", new CultureInfo("es-ES")),
                    EndDate = x.EndDate == null ? null : x.EndDate.GetValueOrDefault().ToString("dd MMM yyyy", new CultureInfo("es-ES"))
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

            var people = _context.ContractDetails.Include(x=>x.People).Where(x => x.BranchesId == id).Select(x=>x.People).Distinct();
            int i = people.Count();
            string res = "";

            foreach (var person in people)
            {
                var contract = person.GetLastContract();
                res += contract.People.CUNI + ";";
                res += contract.People.Document + ";";
                res += contract.People.FirstSurName + ";";
                res += contract.People.SecondSurName + ";";
                res += contract.People.MariedSurName + ";";
                res += contract.People.Names + ";";
                res += contract.People.BirthDate + ";";

                res += contract.Dependency.Cod + ";";
                res += contract.Dependency.Name + ";";

                res += contract.Branches.Abr + ";";
                res += contract.Branches.Name;

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
        [HttpPost]
        [Route("api/Contract/AltaExcel")]
        public async Task<HttpResponseMessage> UploadORExcel()
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
                ContractExcel contractExcel = new ContractExcel(o.excelStream, _context, o.fileName, segId.Id, headerin: 1, sheets: 1);
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
                return BadRequest();
            People person = _context.Person.FirstOrDefault(x => x.CUNI == contract.CUNI.ToString());

            contract.PeopleId = person.Id;
            contract.CUNI = person.CUNI;


            contract.Id = ContractDetail.GetNextId(_context);

            _context.ContractDetails.Add(contract);
            _context.SaveChanges();
            return Created(new Uri(Request.RequestUri + "/" + contract.Id), contract);
        }

        //Bajas
        // POST api/Contract/Baja/5
        [HttpPost]
        [Route("api/Contract/Baja/{id}")]
        public IHttpActionResult Baja(int id, ContractDetail contract)
        {
            ContractDetail contractInDB = _context.ContractDetails.FirstOrDefault(d => d.Id == id);
            // contractInDB.EndDate=DateTime.Now;
            contractInDB.EndDate = contract.EndDate;
            contractInDB.Cause = contract.Cause;
            _context.SaveChanges();
            return Ok(contractInDB);
        }

        // PUT api/Contract/5
        [HttpPut]
        public IHttpActionResult Put(int id, [FromBody]ContractDetail contract)
        {
            if (!ModelState.IsValid)
                return BadRequest();

            ContractDetail contractInDB = _context.ContractDetails.FirstOrDefault(d => d.Id == id);
            if (contractInDB == null)
                return NotFound();

            contractInDB.StartDate = contract.StartDate;
            contractInDB.Dedication = contract.Dedication;
            contractInDB.BranchesId = contract.BranchesId;
            contractInDB.DependencyId = contract.DependencyId;
            contractInDB.PositionsId = contract.PositionsId;
            contractInDB.PositionDescription = contract.PositionDescription;
            contractInDB.Linkage = contract.Linkage;
            contractInDB.AI = contract.AI;

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
