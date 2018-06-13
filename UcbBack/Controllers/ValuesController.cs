using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web;
using System.Web.Http;
using UcbBack.Models;

namespace UcbBack.Controllers
{
    public class ValuesController : ApiController
    {
        private ApplicationDbContext _context;

        public ValuesController()
        {
            _context = new ApplicationDbContext();
        }
        
        // GET api/values
        public List<Pais> Get()
        {
            return _context.Paises.ToList();
        }

        // GET api/values/5
        public People Get(int id)
        {

            var personInDB = _context.Person.FirstOrDefault(d => d.Id == id);
            if (personInDB==null)
                throw new HttpResponseException(HttpStatusCode.NotFound);

            return personInDB;
        }

        // POST api/values
        [HttpPost]
        public IHttpActionResult Post([FromBody]Pais person)
        {
            if(!ModelState.IsValid)
                return BadRequest();

            person.Id = _context.Database.SqlQuery<int>("SELECT \"rrhh_Paises_sqs\".nextval FROM DUMMY;").ToList()[0];
            _context.Paises.Add(person);
            _context.SaveChanges();
            person = (Pais)_context.Entry(person).GetDatabaseValues().ToObject();
            return Created(new Uri(Request.RequestUri + "/" + person.Id), person);
        }

        // PUT api/values/5
        [HttpPut]
        public People Put(int id, [FromBody]People person)
        {
            if (!ModelState.IsValid)
                throw new HttpResponseException(HttpStatusCode.BadRequest);

            var personInDB = _context.Person.First(d => d.Id == id);
            if (personInDB == null)
                throw new HttpResponseException(HttpStatusCode.NotFound);

           // personInDB.NAMES = person.NAMES;
           // personInDB.FIRSTSURNAME = person.FIRSTSURNAME;
           // personInDB.SECONDSURNAME = person.SECONDSURNAME;
           // personInDB.BIRTHDATE = person.BIRTHDATE;

            _context.SaveChanges();
            return person;
        }

        // DELETE api/values/5
        [HttpDelete]
        public void Delete(int id)
        {
            var personInDB = _context.Person.First(d => d.Id == id);
            if (personInDB == null)
                throw new HttpResponseException(HttpStatusCode.NotFound);
            _context.Person.Remove(personInDB);
            _context.SaveChanges();

        }
    }
}
