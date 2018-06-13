using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Helpers;
using System.Web.Http;
using UcbBack.Models;

namespace UcbBack.Controllers
{
    
    public class BranchesController : ApiController
    {
        private ApplicationDbContext _context;    

        public BranchesController()
        {
            _context = new ApplicationDbContext();
        }

        // GET api/Branches
        public IHttpActionResult Get()
        {
                return Ok(_context.Branch.ToList()); 
        }

        // GET api/Branches/5
        public IHttpActionResult Get(int id)
        {
            Branches branchInDB = null;

            branchInDB = _context.Branch.FirstOrDefault(d => d.Id == id);
            
            if (branchInDB == null)
                return NotFound();

            return Ok(branchInDB);
        }

        // POST api/Branches
        [HttpPost]
        public IHttpActionResult Post([FromBody]Branches branch)
        {
            if (!ModelState.IsValid)
                return BadRequest();
            branch.Id = _context.Database.SqlQuery<int>("SELECT \"rrhh_Branches_sqs\".nextval FROM DUMMY;").ToList()[0];

            _context.Branch.Add(branch);
            _context.SaveChanges();
            return Created(new Uri(Request.RequestUri + "/" + branch.Id), branch);
        }

        // PUT api/Branches/5
        [HttpPut]
        public IHttpActionResult Put(int id, [FromBody]Branches branch)
        {
            if (!ModelState.IsValid)
                return BadRequest();

            Branches brachInDB = _context.Branch.FirstOrDefault(d => d.Id == id);
            if (brachInDB == null)
                return NotFound();

            brachInDB.Name = branch.Name;
            brachInDB.Abr = branch.Abr;

            _context.SaveChanges();
            return Ok(brachInDB);
        }

        // DELETE api/Branches/5
        [HttpDelete]
        public IHttpActionResult Delete(int id)
        {
            var brachInDB = _context.Branch.FirstOrDefault(d => d.Id == id);
            if (brachInDB == null)
                return NotFound();
            _context.Branch.Remove(brachInDB);
            _context.SaveChanges();
            return Ok();
        }
    }
}
