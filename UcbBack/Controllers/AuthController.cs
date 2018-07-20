using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Web.Http;
using Newtonsoft.Json.Linq;
using UcbBack.Logic;
using UcbBack.Models;
using UcbBack.Models.Auth;
using System.Data.Entity;

namespace UcbBack.Controllers
{
    public class AuthController : ApiController
    {
        private ApplicationDbContext _context;
        private ValidateToken validator;

        public AuthController()
        {
            _context = new ApplicationDbContext();
            validator = new ValidateToken();
        }

        [HttpGet]
        [Route("api/auth/GetAccess")]
        public IHttpActionResult GetAccess()
        {
            int userid = Int32.Parse(Request.Headers.GetValues("id").First());
            var user = _context.CustomUsers.FirstOrDefault(cu => cu.Id == userid);
            var access = _context.RolshaAccesses.Include(a=>a.Access).Where(a => a.Rolid == user.RolId).Select(x=>new{x.Access.Method,x.Access.Path,x.Access.Description});
            return Ok(access);
        }

        // POST: /api/auth/gettoken/
        [HttpPost]
        [Route("api/auth/GetToken")]
        public IHttpActionResult GetToken([FromBody]JObject credentials)
        {
            if (credentials["username"] == null || credentials["password"] == null)
                return BadRequest();

            string username = credentials["username"].ToString();
            string password = credentials["password"].ToString();

            CustomUser user = _context.CustomUsers.FirstOrDefault(u => u.UserName == username);

            if(user==null)
                return Unauthorized();

            byte[] hashBytes = Convert.FromBase64String(user.Password);
            PasswordHash hash = new PasswordHash(hashBytes);

            if (!hash.Verify(password))
                return Unauthorized();

            user.Token = validator.getToken(user);
            user.TokenCreatedAt = DateTime.Now;
            user.RefreshToken = validator.getRefreshToken(user);
            user.RefreshTokenCreatedAt = DateTime.Now;
            _context.SaveChanges();

            //HttpResponseMessage response = new HttpResponseMessage(HttpStatusCode.OK);
            //response.Headers.Add("Id", user.Id.ToString());
            //response.Headers.Add("Token", user.Token);
            //response.Headers.Add("RefreshToken", user.RefreshToken);
            //return ResponseMessage(response);
            dynamic respose = new JObject();
            respose.Id = user.Id;
            respose.Token = user.Token;
            respose.RefreshToken = user.RefreshToken;
            return Ok(respose);
        }

        [HttpPost]
        [Route("api/auth/AddAccess")]
        public IHttpActionResult AddAccess([FromBody]JObject credentials)
        {
            int rolid = 0;
            int accessid = 0;

            if (!Int32.TryParse(credentials["RolId"].ToString(), out rolid) ||
                !Int32.TryParse(credentials["AccessId"].ToString(), out accessid))
                return BadRequest();

            Rol rol = _context.Rols.FirstOrDefault(r => r.Id == rolid);
            Access access = _context.Accesses.FirstOrDefault(a => a.Id == accessid);

            if (rol == null || access == null)
                return NotFound();

            RolhasAccess rolhasAccess = new RolhasAccess();
            rolhasAccess.Id = _context.Database.SqlQuery<int>("SELECT \"rrhh_RolhasAccess_sqs\".nextval FROM DUMMY;").ToList()[0];
            rolhasAccess.Accessid = accessid;
            rolhasAccess.Rolid = rolid;
            _context.RolshaAccesses.Add(rolhasAccess);
            _context.SaveChanges();
            return Ok();
        }

        // POST: /api/auth/RefreshToken/5
        [HttpPost]
        [Route("api/auth/RefreshToken/{id}")]
        public IHttpActionResult RefreshToken(int id)
        {
            IEnumerable<string> refreshtoken;
            if (!Request.Headers.TryGetValues("RefreshToken", out refreshtoken))
                return BadRequest();

            CustomUser user = _context.CustomUsers.FirstOrDefault(u => u.Id == id && u.RefreshToken == refreshtoken.ElementAt(0));
            if (user == null)
                return Unauthorized();

            user.Token = validator.getToken(user);
            user.TokenCreatedAt = DateTime.Now;
            
            _context.SaveChanges();

            //HttpResponseMessage response = new HttpResponseMessage(HttpStatusCode.OK);
            //response.Headers.Add("Token", user.Token);
            //return ResponseMessage(response);
            dynamic respose = new JObject();
            respose.NewToken = user.Token;
            return Ok(respose);
        }
       
        [HttpGet]
        [Route("api/auth/Logout/{id}")]
        public IHttpActionResult Logout(int id)
        {
            IEnumerable<string> refreshtoken;
            if (!Request.Headers.TryGetValues("RefreshToken",out refreshtoken))
                return BadRequest();

            CustomUser user = _context.CustomUsers.FirstOrDefault(u => u.Id == id && u.RefreshToken == refreshtoken.ElementAt(0));
            if (user == null)
                return Unauthorized();

            user.Token = null;
            user.TokenCreatedAt = null;
            user.RefreshToken = null;
            user.RefreshTokenCreatedAt = null;

            return Ok();
        }
    }
}
