using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using UcbBack.Models;

namespace UcbBack.Controllers
{
    public class DependencyController : ApiController
    {
        private ApplicationDbContext _context;

        public DependencyController()
        {
            _context = new ApplicationDbContext();
        }

        // GET api/Level
        public IHttpActionResult Get()
        {
            return Ok(_context.Dependencies.ToList());
        }

        // GET api/Level/5
        public IHttpActionResult Get(int id)
        {
            Dependency depInDB = null;

            depInDB = _context.Dependencies.FirstOrDefault(d => d.Id == id);

            if (depInDB == null)
                return NotFound();

            return Ok(depInDB);
        }

        // POST api/Level
        [HttpPost]
        public IHttpActionResult Post([FromBody]Dependency dependency)
        {
            if (!ModelState.IsValid)
                return BadRequest();

            dependency.Id = _context.Database.SqlQuery<int>("SELECT \"rrhh_Dependency_sqs\".nextval FROM DUMMY;").ToList()[0];
            _context.Dependencies.Add(dependency);
            _context.SaveChanges();
            return Created(new Uri(Request.RequestUri + "/" + dependency.Id), dependency);
        }

        // PUT api/Level/5
        [HttpPut]
        public IHttpActionResult Put(int id, [FromBody]Dependency dependency)
        {
            if (!ModelState.IsValid)
                return BadRequest();

            Dependency depInDB = _context.Dependencies.FirstOrDefault(d => d.Id == id);
            if (depInDB == null)
                return NotFound();

            depInDB.Cod = dependency.Cod;
            depInDB.Name = dependency.Name;
            depInDB.Parent = dependency.Parent;
            depInDB.OrganizationalUnitId = dependency.OrganizationalUnitId;

            _context.SaveChanges();
            return Ok(depInDB);
        }

        // DELETE api/Level/5
        [HttpDelete]
        public IHttpActionResult Delete(int id)
        {
            var depInDB = _context.Dependencies.FirstOrDefault(d => d.Id == id);
            if (depInDB == null)
                return NotFound();
            _context.Dependencies.Remove(depInDB);
            _context.SaveChanges();
            return Ok();
        }
    }
}
