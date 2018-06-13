using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Web.Mvc;
using System.Web.UI.WebControls;
using Newtonsoft.Json.Linq;
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
        [System.Web.Http.HttpPost]
        public IHttpActionResult Post([FromBody]People person)
        {
            if (!ModelState.IsValid)
                return BadRequest();

            person = validator.CleanName(person);
            person = validator.UcbCode(person);
            //verificar que el ci no exista
            if (_context.Person.FirstOrDefault(p => p.Document == person.Document) != null)
                return BadRequest("El Documento de identidad ya existe.");

            // verificar si existen personas existentes con una similitud mayor al 90%
            var similarities = validator.VerifyExisting(person,0.9f);
            var s = similarities.Count();
            //si existe alguna similitud pedir confirmacion
            if (s > 0)
            {
                string calconftoken = validator.GetConfirmationToken(similarities);
                IEnumerable<string> confirmationToken = new List<string>();
                dynamic response = new JObject();
                //enviar similitudes en la confirmacion
                response.similarities = JToken.FromObject(similarities); ;
                //generar Token de confirmacion con la lista de similitudes
                response.ConfirmationToken = calconftoken;
                //si ya nos estan confirmado se debe enviar el "ConfirmationToken" en los headers,
                //si este token difiere del token calculado volver a pedir confirmacion
                if (Request.Headers.TryGetValues("ConfirmationToken", out confirmationToken))
                    if (calconftoken != confirmationToken.ElementAt(0))
                        return Ok(response);
                //enviar confirmacion
                else return Ok(response);
            }
            
            //si pasa la confirmacion anterior se le asigna un id y se guarda la nueva persona en la BDD
            person.Id = _context.Database.SqlQuery<int>("SELECT \"rrhh_People_sqs\".nextval FROM DUMMY;").ToList()[0];
            
            _context.Person.Add(person);
            _context.SaveChanges();
            return Created(new Uri(Request.RequestUri + "/" + person.Id), person);
        }

        

        // PUT api/People/5
        [System.Web.Http.HttpPut]
        public IHttpActionResult Put(int id, [FromBody]People person)
        {
            if (!ModelState.IsValid)
                return BadRequest();

            People personInDB = _context.Person.FirstOrDefault(d => d.Id == id);
            if (personInDB == null)
                return NotFound();
            //--------------------------REQUIRED COLS--------------------------
            personInDB.TypeDocument = person.TypeDocument;
            personInDB.Document = person.Document;
            personInDB.Ext = person.Ext;
            personInDB.Names = person.Names;
            personInDB.FirstSurName = person.FirstSurName;
            personInDB.SecondSurName = person.SecondSurName;
            personInDB.BirthDate = person.BirthDate;
            personInDB.Gender = person.Gender;
            personInDB.Nationality = person.Nationality;
            //------------------------NON REQUIRED COLS--------------------------
            personInDB.MariedSurName = person.MariedSurName;
            personInDB.PhoneNumber = person.PhoneNumber;
            personInDB.PersonalEmail = person.PersonalEmail;
            personInDB.OfficePhoneNumber = person.OfficePhoneNumber;
            personInDB.OfficePhoneNumberExt = person.OfficePhoneNumberExt;
            personInDB.HomeAddress = person.HomeAddress;
            personInDB.UcbEmail = person.UcbEmail;
            personInDB.AFP = person.AFP;
            personInDB.NUA = person.NUA;
            personInDB.Insurance = person.Insurance;
            personInDB.InsuranceNumber = person.InsuranceNumber;

            _context.SaveChanges();
            return Ok(personInDB);
        }

        // DELETE api/People/5
        [System.Web.Http.HttpDelete]
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
