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
        private ValidateAuth validateauth;

        public AuthController()
        {
            _context = new ApplicationDbContext();
            validator = new ValidateToken();
            validateauth = new ValidateAuth();
        }

        [HttpGet]
        [Route("api/auth/GetMenu")]
        public IHttpActionResult GetMenu()
        {
            int userid;
            IEnumerable<string> headerId;
            if (!Request.Headers.TryGetValues("id", out headerId))
                return BadRequest();
            if (!Int32.TryParse(headerId.FirstOrDefault(), out userid))
                return Unauthorized();

            var user = _context.CustomUsers.FirstOrDefault(cu => cu.Id == userid);
            if (user == null)
                return Unauthorized();

            List<Access> access;
            if (user.RolId == 1)
            {
                access = _context.Accesses.Include(a => a.Resource.Module).Include(a => a.Resource).ToList();
            }
            else
            {
                var rolaccess = _context.RolshaAccesses.Include(a => a.Access).Include(a => a.Access.Resource.Module).Include(a => a.Access.Resource).Where(a => a.Rolid == user.RolId).ToList();
                access = _context.Accesses.Include(a => a.Resource.Module).Include(a => a.Resource).ToList();
                access = access.Where(a => rolaccess.ToList().Where(r => r.Accessid == a.Id).Count() > 0).ToList();
            } 
            List<dynamic> res = new List<dynamic>();
            var listModules = access.Select(a=>a.Resource.Module).Distinct();
            var listResources = access.Select(a=>a.Resource).Distinct();
            foreach (var module in listModules)
            {
                List<dynamic> children = new List<dynamic>();
                foreach (var child in listResources.Where(c=>c.ModuleId==module.Id))
                {
                    var listmethods = access.Where(a => a.ResourceId == child.Id).Select(a=>a.Method).Distinct();
                    dynamic c = new JObject();
                    c.name = child.Name;
                    c.path = child.Path;
                    c.methods = JArray.FromObject(listmethods.ToArray());
                    children.Add(c);
                }

                dynamic r = new JObject();
                r.name = module.Name;
                r.icon = module.Icon;
                r.collapsed = true;
                r.children = JArray.FromObject(children.ToArray());
                res.Add(r);
            }
            return Ok(res);
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

            CustomUser user = _context.CustomUsers.Include(u => u.Rol).Include(u => u.Rol.Resource).FirstOrDefault(u => u.UserName == username);

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
            respose.ExpiresIn = validateauth.tokenLife;
            respose.RefreshExpiresIn = validateauth.refeshtokenLife;
            respose.AccessDefault = user.Rol.Resource.Path;
            return Ok(respose);
        }

        // POST: /api/auth/RefreshToken/
        [HttpPost]
        [Route("api/auth/RefreshToken/")]
        public IHttpActionResult RefreshToken(JObject data)
        {

            IEnumerable<string> idlist;
            if (!Request.Headers.TryGetValues("id", out idlist))
                return BadRequest();
            if (data["RefreshToken"] == null)
                return BadRequest();

            int userid = 0;
            if (!Int32.TryParse(idlist.First(), out userid))
                return Unauthorized();
            string rt = data["RefreshToken"].ToString();
            CustomUser user = _context.CustomUsers.FirstOrDefault(u => u.Id == userid && u.RefreshToken == rt);
            if (user == null)
                return Unauthorized();
            if (user.RefreshTokenCreatedAt == null)
                return Unauthorized();

            int seconds = (int)DateTime.Now.Subtract(user.RefreshTokenCreatedAt.Value).TotalSeconds;

            if (seconds > validateauth.refeshtokenLife)
                return Unauthorized();

            user.Token = validator.getToken(user);
            user.TokenCreatedAt = DateTime.Now;
            
            _context.SaveChanges();

            dynamic respose = new JObject();
            respose.Token = user.Token;
            respose.ExpiresIn = validateauth.tokenLife;

            return Ok(respose);
        }
       
        [HttpGet]
        [Route("api/auth/Logout/")]
        public IHttpActionResult Logout()
        {
            IEnumerable<string> tokenlist;
            IEnumerable<string> idlist;
            if (!Request.Headers.TryGetValues("token", out tokenlist) || !Request.Headers.TryGetValues("id", out idlist))
                return Unauthorized();
            int userid = 0;

            if (!Int32.TryParse(idlist.First(), out userid))
                return Unauthorized();

            string token = tokenlist.First();
            CustomUser user = _context.CustomUsers.FirstOrDefault(u => u.Id == userid && u.Token == token);
            if (user == null)
                return Unauthorized();

            user.Token = null;
            user.TokenCreatedAt = null;
            user.RefreshToken = null;
            user.RefreshTokenCreatedAt = null;
            _context.SaveChanges();

            return Ok();
        }
    }
}
