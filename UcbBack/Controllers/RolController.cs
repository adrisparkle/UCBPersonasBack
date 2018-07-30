using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using Newtonsoft.Json.Linq;
using UcbBack.Logic;
using UcbBack.Models;
using UcbBack.Models.Auth;
using System.Data.Entity;

namespace UcbBack.Controllers
{
    public class RolController : ApiController
    {
        private ApplicationDbContext _context;
        private ValidatePerson validator;

        public RolController()
        {
            _context = new ApplicationDbContext();
            validator = new ValidatePerson(_context);
        }

        // GET api/Rol
        public IHttpActionResult Get()
        {
            return Ok(_context.Rols.Include("Resource").Select(r => new { r.Id, r.Name, r.Level, r.ResourceId, Resource=r.Resource.Name }).ToList());
        }

        // GET api/Rol/5
        public IHttpActionResult Get(int id)
        {
            Rol rolInDB = null;

            rolInDB = _context.Rols.FirstOrDefault(d => d.Id == id);

            if (rolInDB == null)
                return NotFound();

            return Ok(rolInDB);
        }

        // POST api/Rol
        [HttpPost]
        public IHttpActionResult Post([FromBody]Rol rol)
        {
            if (!ModelState.IsValid)
                return BadRequest();
            rol.Id = _context.Database.SqlQuery<int>("SELECT \"rrhh_Rol_sqs\".nextval FROM DUMMY;").ToList()[0];
            _context.Rols.Add(rol);
            _context.SaveChanges();
            return Created(new Uri(Request.RequestUri + "/" + rol.Id), rol);
        }



        // PUT api/Rol/5
        [HttpPut]
        public IHttpActionResult Put(int id, [FromBody]Rol rol)
        {
            if (!ModelState.IsValid)
                return BadRequest();

            Rol rolInDB = _context.Rols.FirstOrDefault(d => d.Id == id);
            if (rolInDB == null)
                return NotFound();

            rolInDB.Name = rol.Name;
            rolInDB.Level = rol.Level;
            rolInDB.ResourceId = rol.ResourceId;
            _context.SaveChanges();

            return Ok(rolInDB);
        }

        // DELETE api/Rol/5
        [HttpDelete]
        public IHttpActionResult Delete(int id)
        {
            var rolInDB = _context.Rols.FirstOrDefault(d => d.Id == id);

            if (rolInDB == null)
                return NotFound();

            _context.Rols.Remove(rolInDB);
            _context.SaveChanges();

            return Ok();
        }

        [HttpGet]
        [Route("api/rol/GetAccess/{id}")]
        public IHttpActionResult GetAccess(int id)
        {
           var rha = _context.RolshaAccesses.Include(r=>r.Access).Where(r => r.Rolid == id).
                 Select(r => new
                 {
                     r.Access.Id,
                     r.Access.Method,
                     r.Access.Description,
                     r.Access.Path,
                     r.Access.Public
                 }).ToList();

            return Ok(rha);
        }

        [HttpPost]
        [Route("api/rol/AddAccess/{id}")]
        public IHttpActionResult AddAccess(int id,[FromBody]JObject credentials)
        {
            int accessid = 0;
            if (credentials["AccessId"] == null)
                return BadRequest();

            if (!Int32.TryParse(credentials["AccessId"].ToString(), out accessid))
                return BadRequest();

            Rol rol = _context.Rols.FirstOrDefault(r => r.Id == id);
            Access access = _context.Accesses.FirstOrDefault(a => a.Id == accessid);

            if (rol == null || access == null)
                return NotFound();

            RolhasAccess rolhasAccess = new RolhasAccess();
            rolhasAccess.Id = _context.Database.SqlQuery<int>("SELECT \"rrhh_RolhasAccess_sqs\".nextval FROM DUMMY;").ToList()[0];
            rolhasAccess.Accessid = accessid;
            rolhasAccess.Rolid = id;
            _context.RolshaAccesses.Add(rolhasAccess);
            _context.SaveChanges();

            return Ok();
        }
    }
}
