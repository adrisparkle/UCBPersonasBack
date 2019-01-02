using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using UcbBack.Models;
using System.Data.Entity;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json.Linq;
using UcbBack.Logic;

namespace UcbBack.Controllers
{
    public class PositionsController : ApiController
    {
        private ApplicationDbContext _context;
        private ValidateAuth auth;

        public PositionsController()
        {
            _context = new ApplicationDbContext();
            auth = new ValidateAuth();
        }

        // GET api/Positions
        public IHttpActionResult Get()
        {
            var all = from pos in _context.Position.Include(x => x.Level)
                join brs in _context.BranchhasPositions on pos.Id equals brs.PositionId
                where brs.Enabled
                select new
                {
                    pos.Id,
                    pos.Name,
                    pos.Level.Cod,
                    pos.Level.Category,
                    brs.BranchesId 
                };
            var user = auth.getUser(Request);
            var filtered = auth.filerByRegional(all.AsQueryable(), user);
            List<dynamic> res = new List<dynamic>();
            foreach (var p in filtered)
            {
                dynamic r = new JObject();
                r.Id = p.Id;
                r.Name = p.Name;
                r.Cod = p.Cod;
                r.Category = p.Category;
                if (!res.Any(x=>x.Id==r.Id))
                    res.Add(r);
            }

            var xs = res.ToList().Distinct();
            return Ok(xs); 
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
            position.Id = Positions.GetNextId(_context);
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
