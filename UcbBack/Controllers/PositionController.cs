using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using UcbBack.Models;
using System.Data.Entity;

namespace UcbBack.Controllers
{
    public class PositionsController : ApiController
    {
        private ApplicationDbContext _context;    

        public PositionsController()
        {
            _context = new ApplicationDbContext();
        }

        // GET api/Positions
        public IHttpActionResult Get()
        {
            var poslist = _context.Position.Include(p=>p.Level).Include(i=>i.Branches).ToList().Select(x=>new{x.Id,x.Name,x.Level.Cod,x.Level.Category,Branch = x.Branches.Name});
            return Ok(poslist); 
        }

        // GET api/Positions/5
        public IHttpActionResult Get(int id)
        {
            Positions positionInDB = null;

            positionInDB = _context.Position.FirstOrDefault(d => d.Id == id);

            if (positionInDB == null)
                return NotFound();

            return Ok(positionInDB);
        }

        // POST api/Positions
        [HttpPost]
        public IHttpActionResult Post([FromBody]Positions position)
        {
            if (!ModelState.IsValid)
                return BadRequest();
            position.Id = _context.Database.SqlQuery<int>("SELECT \"rrhh_Position_sqs\".nextval FROM DUMMY;").ToList()[0];
            _context.Position.Add(position);
            _context.SaveChanges();
            return Created(new Uri(Request.RequestUri + "/" + position.Id), position);
        }

        // PUT api/Positions/5
        [HttpPut]
        public IHttpActionResult Put(int id, [FromBody]Positions position)
        {
            if (!ModelState.IsValid)
                return BadRequest();

            Positions positionInDB = _context.Position.FirstOrDefault(d => d.Id == id);
            if (positionInDB == null)
                return NotFound();

            positionInDB.Name = position.Name;
            positionInDB.BranchesId = position.BranchesId;
            positionInDB.LevelId = position.LevelId;

            _context.SaveChanges();
            return Ok(positionInDB);
        }

        // DELETE api/Positions/5
        [HttpDelete]
        public IHttpActionResult Delete(int id)
        {
            var positionInDB = _context.Position.FirstOrDefault(d => d.Id == id);
            if (positionInDB == null)
                return NotFound();
            _context.Position.Remove(positionInDB);
            _context.SaveChanges();
            return Ok();
        }
    }
}
