using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using UcbBack.Logic;
using UcbBack.Models;
using UcbBack.Models.Auth;

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
            return Ok(_context.Rols.ToList());
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
    }
}
