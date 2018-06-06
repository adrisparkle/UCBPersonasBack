using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Web.UI.WebControls;
using UcbBack.Logic;
using UcbBack.Models;

namespace UcbBack.Controllers
{
    public class PeopleController : ApiController
    {
        private ApplicationDbContext _context;
        private ValidatePerson validator;

        public PeopleController()
        {
            _context = new ApplicationDbContext();
            validator = new ValidatePerson(_context);
        }

        // GET api/People
        public IHttpActionResult Get()
        {
            return Ok(_context.Person.ToList()); 
        }

        // GET api/People/5
        public IHttpActionResult Get(int id)
        {
            People personInDB = null;

            personInDB = _context.Person.FirstOrDefault(d => d.Id == id);

            if (personInDB == null)
                return NotFound();

            return Ok(personInDB);
        }

        // POST api/People
        [HttpPost]
        public IHttpActionResult Post([FromBody]People person)
        {
            if (!ModelState.IsValid)
                return BadRequest();
            person = validator.UcbCode(person);
            _context.Person.Add(person);
            _context.SaveChanges();
            return Created(new Uri(Request.RequestUri + "/" + person.Id), person);
        }

        

        // PUT api/People/5
        [HttpPut]
        public IHttpActionResult Put(int id, [FromBody]People person)
        {
            if (!ModelState.IsValid)
                return BadRequest();

            People personInDB = _context.Person.FirstOrDefault(d => d.Id == id);
            if (personInDB == null)
                return NotFound();
            //--------------------------REQUIRED COLS--------------------------
            personInDB.TYPE_DOCUMENT = person.TYPE_DOCUMENT;
            personInDB.DOCUMENTO = person.DOCUMENTO;
            personInDB.ISSUED = person.ISSUED;
            personInDB.NAMES = person.NAMES;
            personInDB.FIRSTSURNAME = person.FIRSTSURNAME;
            personInDB.SECONDSURNAME = person.SECONDSURNAME;
            personInDB.BIRTHDATE = person.BIRTHDATE;
            personInDB.BIRTHDATE = person.BIRTHDATE;
            personInDB.GENDER = person.GENDER;
            personInDB.NATIONALITY = person.NATIONALITY;
            //------------------------NON REQUIRED COLS--------------------------
            personInDB.MARIEDSURNAME = person.MARIEDSURNAME;
            personInDB.PHONENUMBER = person.PHONENUMBER;
            personInDB.PERSONALEMAIL = person.PERSONALEMAIL;
            personInDB.OFFICEPHONENUMBER = person.OFFICEPHONENUMBER;
            personInDB.HOMEADDRESS = person.HOMEADDRESS;
            personInDB.MARITALSTATUS = person.MARITALSTATUS;
            personInDB.UCBMAIL = person.UCBMAIL;

            _context.SaveChanges();
            return Ok(personInDB);
        }

        // DELETE api/People/5
        [HttpDelete]
        public IHttpActionResult Delete(int id)
        {
            var personInDB = _context.Person.FirstOrDefault(d => d.Id == id);
            if (personInDB == null)
                return NotFound();
            _context.Person.Remove(personInDB);
            _context.SaveChanges();
            return Ok();
        }
    }
}
