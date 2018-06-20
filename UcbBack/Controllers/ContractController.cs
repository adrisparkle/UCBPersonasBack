using System;
using System.Collections.Generic;
using System.Data;
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
            return Ok(_context.Contracts.ToList());
        }

        // GET api/Contract/5
        public IHttpActionResult Get(int id)
        {
            Contract contractInDB = null;

            contractInDB = _context.Contracts.FirstOrDefault(d => d.Id == id);

            if (contractInDB == null)
                return NotFound();

            return Ok(contractInDB);
        }

        [HttpPost]
        [Route("api/Contract/AltaExcel")]
        public async Task<HttpResponseMessage> AltasExcel()
        {
            
            var response = new HttpResponseMessage();
            try
            {
                var req = await Request.Content.ReadAsMultipartAsync();
                foreach (HttpContent contentPart in req.Contents)
                {
                    Stream stream = await contentPart.ReadAsStreamAsync();
                    var contentDisposition = contentPart.Headers.ContentDisposition;
                    string fileName = String.IsNullOrEmpty(contentDisposition.FileName)
                        ? ""
                        : contentDisposition.FileName.Trim('"');
                    var mediaType = contentPart.Headers.ContentType == null ? "" :
                        String.IsNullOrEmpty(contentPart.Headers.ContentType.MediaType) ? "" :
                        contentPart.Headers.ContentType.MediaType;
                    ValidateContractsFile contractExcel = new ValidateContractsFile(stream, fileName,_context);

                    return contractExcel.toResponse();
                }
                response.StatusCode = HttpStatusCode.BadRequest;
                response.Content = new StringContent("Por favor enviar un archivo en formato excel (.xls, .xslx)");
                return response;
            }
            catch (System.ArgumentException e)
            {
                response.StatusCode = HttpStatusCode.BadRequest;
                response.Content = new StringContent("Por favor enviar un archivo en formato excel (.xls, .xslx)");
                return response;
            }
            // Scan the Multiple Parts 
            
        }

        //altas
        // POST api/Contract
        [HttpPost]
        [Route("api/Contract/Alta")]
        public IHttpActionResult Post([FromBody]Contract contract)
        {
            if (!ModelState.IsValid)
                return BadRequest();
            contract.Id = _context.Database.SqlQuery<int>("SELECT \"rrhh_Contract_sqs\".nextval FROM DUMMY;").ToList()[0];
            _context.Contracts.Add(contract);
            _context.SaveChanges();
            return Created(new Uri(Request.RequestUri + "/" + contract.Id), contract);
        }

        [HttpPost]
        [Route("api/Contract/Baja/{id}")]
        public IHttpActionResult Baja(int id,[FromBody]Contract contract)
        {
            if (!ModelState.IsValid)
                return BadRequest();
            contract.Id = _context.Database.SqlQuery<int>("SELECT \"rrhh_Contract_sqs\".nextval FROM DUMMY;").ToList()[0];
            _context.Contracts.Add(contract);
            _context.SaveChanges();
            return Created(new Uri(Request.RequestUri + "/" + contract.Id), contract);
        }

        [HttpPost]
        [Route("api/Contract/Modificaciones/{id}")]
        public IHttpActionResult Modificaciones(int id,[FromBody]Contract contract)
        {
            if (!ModelState.IsValid)
                return BadRequest();
            contract.Id = _context.Database.SqlQuery<int>("SELECT \"rrhh_Contract_sqs\".nextval FROM DUMMY;").ToList()[0];
            _context.Contracts.Add(contract);
            _context.SaveChanges();
            return Created(new Uri(Request.RequestUri + "/" + contract.Id), contract);
        }

        // PUT api/Contract/5
        [HttpPut]
        public IHttpActionResult Put(int id, [FromBody]Contract contract)
        {
            if (!ModelState.IsValid)
                return BadRequest();

            Contract contractInDB = _context.Contracts.FirstOrDefault(d => d.Id == id);
            if (contractInDB == null)
                return NotFound();

           // contractInDB.Cod = contract.Cod;
           // contractInDB.Category = contract.Category;

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
