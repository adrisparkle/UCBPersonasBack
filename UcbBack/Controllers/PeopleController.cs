using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Web.UI.WebControls;
using Newtonsoft.Json.Linq;
using UcbBack.Logic;
using UcbBack.Models;
using System.Data.Entity;


namespace UcbBack.Controllers
{
    public class PeopleController : ApiController
    {
        private ApplicationDbContext _context;
        private ValidatePerson validator;
        private ADClass activeDirectory;

        public PeopleController()
        {
            _context = new ApplicationDbContext();
            validator = new ValidatePerson(_context);
            activeDirectory = new ADClass();
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

        [HttpGet]
        [Route("api/addalltoAD")]
        public IHttpActionResult addalltoAD()
        {
            /*List<string> branches = _context.Branch.Select(x => x.ADGroupName).ToList();
            List<string> roles = _context.Rols.Select(x => x.ADGroupName).ToList();

            foreach (var branch in branches)
            {
                activeDirectory.createGroup(branch);
            }

            foreach (var rol in roles)
            {
                activeDirectory.createGroup(rol);
            }*/


            /*DateTime date = new DateTime(2018,4,1);
            //todo run for all people   solo fata pasar fechas de corte
            List<People> person = _context.ContractDetails.Include(x => x.People).Include(x=>x.Positions).
                Where(y => y.StartDate < date 
                           && (y.EndDate > date || y.EndDate == null) 
                           && y.Positions.Name != "DOCENTE TIEMPO HORARIO").Select(x => x.People).ToList();

            List<string> asd = new List<string>();
            foreach (var pe in person)
            {
                //var tt = activeDirectory.findUser(pe);
                activeDirectory.addUser(pe);
                _context.SaveChanges();
            }*/
            //activeDirectory.createGroup("");
            var yoo = _context.Person.FirstOrDefault(x => x.CUNI == "AMG680422");
            activeDirectory.AddUserToGroup(_context.CustomUsers.FirstOrDefault(x => x.PeopleId == yoo.Id).UserPrincipalName, "Personas.SALOMON");
            var yo = activeDirectory.findUser(_context.Person.FirstOrDefault(x => x.CUNI == "AMG680422"));
            
            return Ok(yo.UserPrincipalName+" - "+yo.DisplayName);
        }
        // POST api/People
        [HttpPost]
        public IHttpActionResult Post([FromBody]People person)
        {
            if (!ModelState.IsValid)
                return BadRequest(ModelState);

            person = validator.CleanName(person);
            person = validator.UcbCode(person);
            //verificar que el ci no exista
            if (_context.Person.FirstOrDefault(p => p.Document == person.Document) != null)
                return BadRequest("El Documento de identidad ya existe.");

            // verificar si existen personas existentes con una similitud mayor al 90%
            var similarities = validator.VerifyExisting(person,0.95f);
            var s = similarities.Count();
            //si existe alguna similitud pedir confirmacion
            /*if (s > 0)
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
            }*/
            
            //si pasa la confirmacion anterior se le asigna un id y se guarda la nueva persona en la BDD
            person.Id = _context.Database.SqlQuery<int>("SELECT ADMNALRRHH.\"rrhh_People_sqs\".nextval FROM DUMMY;").ToList()[0];
            
            _context.Person.Add(person);
            _context.SaveChanges();
            // activeDirectory.addUser(person);

            return Created(new Uri(Request.RequestUri + "/" + person.Id), person);
        }

        

        // PUT api/People/5
        [HttpPut]
        public IHttpActionResult Put(int id, [FromBody]People person)
        {
            if (!ModelState.IsValid)
                return BadRequest(ModelState);

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
    }
}
