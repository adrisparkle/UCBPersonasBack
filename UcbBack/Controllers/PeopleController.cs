using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Web.UI.WebControls;
using Newtonsoft.Json.Linq;
using UcbBack.Logic;
using UcbBack.Models;
using System.Data.Entity;
using System.IO;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using DocumentFormat.OpenXml.Wordprocessing;
using UcbBack.Logic.B1;


namespace UcbBack.Controllers
{
    public class PeopleController : ApiController
    {
        private ApplicationDbContext _context;
        private ValidatePerson validator;
        private ValidateAuth auth;
        private ADClass activeDirectory;

        public PeopleController()
        {
            _context = new ApplicationDbContext();
            validator = new ValidatePerson(_context);
            auth = new ValidateAuth();
            activeDirectory = new ADClass();
        }

        // GET api/People
        public IHttpActionResult Get()
        {
            return Ok(_context.Person.ToList()); 
        }

        
        // GET api/People/5
        [System.Web.Http.HttpGet]
        [System.Web.Http.Route("api/people/Query")]
        public IHttpActionResult Query([FromUri] string by, [FromUri] string value)
        {
            //todo cleanText of "Value"

            People person = null;
            switch (by)
            {
                case "CUNI":
                    person = _context.Person.FirstOrDefault(x=>x.CUNI==value);
                    break;
                case "Documento":
                    string no_start_zeros = value.Replace("E-","").TrimStart('0');
                    person = _context.Person.ToList().FirstOrDefault(x => x.Document.Replace("E-", "").TrimStart('0') == value.Replace("E-", "").TrimStart('0'));
                    break;
                case "FullName":
                    person = _context.Person.ToList().FirstOrDefault(x => x.GetFullName() == value);
                    break;
                default:
                    return BadRequest();
            }

            if (person == null)
                return NotFound();

            dynamic res = new JObject();
            res.Id = person.Id;
            res.CUNI = person.CUNI;
            res.Document = person.Document;
            res.FullName = person.GetFullName();
            res.contract = person.GetLastContract(_context, DateTime.Now) != null;

            return Ok(res);
        }

        [System.Web.Http.NonAction]
        public async Task<System.Dynamic.ExpandoObject> HttpContentToVariables(MultipartMemoryStreamProvider req)
        {
            dynamic res = new System.Dynamic.ExpandoObject();
            foreach (HttpContent contentPart in req.Contents)
            {
                var contentDisposition = contentPart.Headers.ContentDisposition;
                string varname = contentDisposition.Name;

                if (varname == "\"file\"")
                {
                    Stream stream = await contentPart.ReadAsStreamAsync();
                    res.fileName = String.IsNullOrEmpty(contentDisposition.FileName) ? "" : contentDisposition.FileName.Trim('"');
                    res.excelStream = stream;
                }
            }
            return res;
        }


        [System.Web.Http.HttpPost]
        [System.Web.Http.Route("api/people/UploadCI/{id}")]
        public HttpResponseMessage Index(int id)
        {
            if (!Request.Content.IsMimeMultipartContent("form-data"))
                return Request.CreateResponse(HttpStatusCode.InternalServerError);
            var request = HttpContext.Current.Request;
            bool SubmittedFile = (request.Files.AllKeys.Length > 0);

            var response = new HttpResponseMessage();
            var file = request.Files[0];

            var person = _context.Person.FirstOrDefault(x => x.Id == id);
            if (person == null)
            {
                response.StatusCode = HttpStatusCode.NotFound;
                return response;
            }

            if (SubmittedFile)
            {
                try
                {
                    var type = file.ContentType.ToLower();
                    if (type != "image/jpg" &&
                        type != "image/jpeg" &&
                        type != "image/pjpeg" &&
                        type != "image/gif" &&
                        type != "image/x-png" &&
                        type != "image/png")
                    {
                        response.StatusCode = HttpStatusCode.BadRequest;
                        response.Content = new StringContent("Invalid File");
                        return response;
                    }
                    string path = Path.Combine(HttpContext.Current.Server.MapPath("~/Images/PeopleDocuments"),
                        (person.CUNI + Path.GetExtension(file.FileName)));
                    file.SaveAs(path);

                    person.DocPath = path;
                    _context.SaveChanges();
                    response.StatusCode = HttpStatusCode.OK;
                    response.Content = new StringContent(path);
                    return response;

                }
                catch (Exception ex)
                {
                    response.StatusCode = HttpStatusCode.BadRequest;
                    response.Content = new StringContent("Invalid File"+ex.Message);
                    return response;
                }
            }
            else
            {
                response.StatusCode = HttpStatusCode.BadRequest;
                response.Content = new StringContent("No File");
                return response;
            }
        }

        [System.Web.Http.HttpGet]
        [System.Web.Http.Route("api/people/Contracts/{id}")]
        public IHttpActionResult GetContracts(int id, [FromUri] string now = "NO")
        {
            var contracts = _context.ContractDetails
                .Include(x => x.Dependency)
                .Include(x => x.Positions)
                .Where(x => x.PeopleId == id)
                .ToList()
                .Where(x=> now == "NO" || (x.EndDate == null || x.EndDate.Value > DateTime.Now))
                .OrderByDescending(x => x.EndDate == null ? DateTime.MaxValue : DateTime.MinValue)
                .ThenByDescending(x => x.StartDate)
                .Select(x => new
                {
                    x.Id,
                    x.Dependency.Cod,
                    Dependency = x.Dependency.Name,
                    Positions = x.Positions.Name,
                    //StartDate = x.StartDate.ToString("dd-MM-yyyy"),
                    //EndDate = x.EndDate==null?null:x.EndDate.Value.ToString("dd-MM-yyyy")
                    StartDate = x.StartDate.ToString("MM-dd-yyyy"),
                    EndDate = x.EndDate == null ? null : x.EndDate.Value.ToString("MM-dd-yyyy")
                });
                
            return Ok(contracts);
        }

        // GET api/People/5
        [System.Web.Http.HttpGet]
        [System.Web.Http.Route("api/people/{id}")]
        public IHttpActionResult Get(int id, [FromUri] string by = "Id")
        {

            People personInDB = null;
            switch (by)
            {
                case "Id":
                    personInDB = _context.Person.FirstOrDefault(d => d.Id == id);
                    break;
                case "Contract":
                    var con = _context.ContractDetails.Include(x=>x.People).FirstOrDefault(d => d.Id == id);
                    personInDB = con == null ? null : con.People;
                    break;
            }


            if (personInDB == null)
                return NotFound();

            dynamic res = new JObject();
            res.Id = personInDB.Id;
            res.CUNI = personInDB.CUNI;
            res.Document = personInDB.Document;
            res.TypeDocument = personInDB.TypeDocument;
            res.Ext = personInDB.Ext;
            res.FullName = personInDB.GetFullName();
            res.FirstSurName = personInDB.FirstSurName;
            res.SecondSurName = personInDB.SecondSurName;
            res.Names = personInDB.Names;
            res.MariedSurName = personInDB.MariedSurName == null ? "" : personInDB.MariedSurName;
            res.UseMariedSurName = personInDB.UseMariedSurName;
            res.UseSecondSurName = personInDB.UseSecondSurName;
            var c = personInDB.GetLastContract(_context, date:DateTime.Now);
            res.Contract = c != null;
            res.ContractId = c == null ? (dynamic) "" : c.Id;
            res.PositionsId = c == null ? (dynamic) "" : c.Positions.Id;
            res.Positions = c == null ? "" : c.Positions.Name;
            res.PositionDescription = c == null ? "" : c.PositionDescription;
            res.AI = c == null ? false : c.AI;
            res.Dedication = c == null ? "" : c.Dedication;
            res.Linkage = c == null ? "" : c.Link.Value;
            res.DependencyId = c == null ? (dynamic) "" : c.Dependency.Id;
            res.Dependency = c == null ? "" : c.Dependency.Name;
            res.Branches = c == null ? null : _context.Branch.FirstOrDefault(x => x.Id == c.Dependency.BranchesId).Name;
            res.StartDate = c == null ? (dynamic)"" : c.StartDate.ToString("MM/dd/yyyy");
            res.EndDate = c == null ? (dynamic)"" : c.EndDate == null ? "" : c.EndDate.Value.ToString("MM/dd/yyyy");
            res.Gender = personInDB.Gender;
            res.BirthDate = personInDB.BirthDate.ToString("MM/dd/yyyy");
            res.Nationality = personInDB.Nationality;
            res.AFP = personInDB.AFP;
            res.NUA = personInDB.NUA;

            var u = _context.CustomUsers.FirstOrDefault(x => x.PeopleId == personInDB.Id);
            res.UserName = u == null ? "" : u.UserPrincipalName;

            return Ok(res);
        }

        [System.Web.Http.HttpGet]
        [System.Web.Http.Route("api/people/manager/{id}")]
        public IHttpActionResult manager(int id)
        {
            People personInDB = _context.Person.FirstOrDefault(d => d.Id == id);
            if (personInDB == null)
                return NotFound();
            return Ok(personInDB.GetLastManager());
        }

        [System.Web.Http.HttpGet]
        [System.Web.Http.Route("api/addalltoAD")]
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

            List<string> palabras = new List<string>(new string []
            {
                "aula",
                "libro",
                "lapiz",
                "papel",
                "folder",
                "lentes"
            });

        DateTime date = new DateTime(2017,1,1);
            //todo run for all people   solo fata pasar fechas de corte
            /*var usr = _context.CustomUsers.Select(x => x.PeopleId).ToList();*/
            List<People> person = _context.ContractDetails.Include(x => x.People).Include(x=>x.Positions).
                Where(y => (y.EndDate > date || y.EndDate == null)
                    ).Select(x => x.People).Distinct().ToList();

            B1Connection b1 = B1Connection.Instance();
            var usr = auth.getUser(Request);
            int i = 0;
            foreach (var p in person)
            {
                i++;
                var X = b1.AddOrUpdatePerson(usr.Id,p);
                if (X.Contains("ERROR"))
                {
                    X = "";
                }
            }

            /*Random rnd = new Random();
            foreach (var pe in person)
            {
                //var tt = activeDirectory.findUser(pe);
                string pass = palabras[rnd.Next(6)];
                while (pass.Length<8)
                {
                    pass += rnd.Next(10);
                }

                var ex = _context.CustomUsers.FirstOrDefault(x => x.PeopleId == pe.Id);
                if (ex == null)
                {
                    activeDirectory.addUser(pe, pass);
                    _context.SaveChanges();
                    var account = _context.CustomUsers.FirstOrDefault(x => x.PeopleId == pe.Id);
                    account.AutoGenPass = pass;
                    _context.SaveChanges();
                }
                
            }*/
            //activeDirectory.createGroup("");
            //var yoo = _context.Person.FirstOrDefault(x => x.CUNI == "RFA940908");
           // activeDirectory.AddUserToGroup(_context.CustomUsers.FirstOrDefault(x => x.PeopleId == yoo.Id).UserPrincipalName, "Personas.Admin");
            //var yo = activeDirectory.findUser(_context.Person.FirstOrDefault(x => x.CUNI == "AMG680422"));
            
            return Ok();
        }
        // POST api/People
        [System.Web.Http.HttpPost]
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
            person.Id = People.GetNextId(_context);
            
            _context.Person.Add(person);
            _context.SaveChanges();
            // activeDirectory.addUser(person);

            dynamic res = new JObject();
            res.Id = person.Id;
            res.CUNI = person.CUNI;
            res.Document = person.Document;
            res.FullName = person.GetFullName();

            return Created(new Uri(Request.RequestUri + "/" + person.Id), res);
        }

        

        // PUT api/People/5
        [System.Web.Http.HttpPut]
        [System.Web.Http.Route("api/people/{id}")]
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
            personInDB.UseMariedSurName = person.UseMariedSurName;
            personInDB.UseSecondSurName = person.UseSecondSurName;
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
