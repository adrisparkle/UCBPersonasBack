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


namespace UcbBack.Controllers
{
    public class ContractController : ApiController
    {
         private ApplicationDbContext _context;

        public ContractController()
        {
            _context = new ApplicationDbContext();
        }

        // GET api/Contract
        public IHttpActionResult Get()
        {
            var contplist = _context.ContractDetails.Include(p => p.Branches).Include(p => p.Dependency).Include(p => p.Positions).Include(p => p.People).ToList().Select(x => new { x.Id, x.People.CUNI, x.People.Document, x.People.FirstSurName, x.People.SecondSurName, x.People.Names, Dependency = x.Dependency.Name, Branches = x.Branches.Abr, Positions=x.Positions.Name, x.Dedication,x.Linkage,x.StartDate,x.EndDate }).OrderBy(x => x.Id);
            return Ok(contplist);
        }

        // GET api/Contract/5
        public IHttpActionResult Get(int id)
        {
            ContractDetail contractInDB = null;

            contractInDB = _context.ContractDetails.FirstOrDefault(d => d.Id == id);

            if (contractInDB == null)
                return NotFound();

            return Ok(contractInDB);
        }
        [HttpGet]
        [Route("api/Contract/GetPersonContract")]
        public IHttpActionResult GetPersonContract(int id)
        {
            List<ContractDetail> contractInDB = null;

            contractInDB = _context.ContractDetails.Where(d => d.People.Id == id).ToList();

            if (contractInDB == null)
                return NotFound();

            return Ok(contractInDB);
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
                else if (varname == "\"uploadfile\"")
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

                if (o.segmentoOrigen == null)
                {
                    response.StatusCode = HttpStatusCode.BadRequest;
                    response.Content = new StringContent("Debe enviar segmentoOrigen");
                    return response;
                }

                string name = o.segmentoOrigen.ToString();
                var segId = _context.Branch.FirstOrDefault(b => b.Name == "");
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
        // POST api/Contract
        [HttpPost]
        [Route("api/Contract/Alta")]
        public IHttpActionResult Post([FromBody]ContractDetail contract)
        {
            if (!ModelState.IsValid)
                return BadRequest();
            contract.Id = _context.Database.SqlQuery<int>("SELECT ADMNALRRHH.\"rrhh_ContractDetail_sqs\".nextval FROM DUMMY;").ToList()[0];
            _context.ContractDetails.Add(contract);
            _context.SaveChanges();
            return Created(new Uri(Request.RequestUri + "/" + contract.Id), contract);
        }

        [HttpPost]
        [Route("api/Contract/Baja/{id}")]
        public IHttpActionResult Baja(int id)
        {
            ContractDetail contractInDB = _context.ContractDetails.FirstOrDefault(d => d.Id == id);
            contractInDB.EndDate=DateTime.Now;
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
            _context.Contracts.Remove(contractInDB);
            _context.SaveChanges();
            return Ok();
        }
    }
}
