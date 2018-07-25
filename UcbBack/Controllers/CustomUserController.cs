using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Web.Http;
using Newtonsoft.Json.Linq;
using UcbBack.Logic;
using UcbBack.Models;
using UcbBack.Models.Auth;

namespace UcbBack.Controllers
{
    public class CustomUserController : ApiController
    {
        private ApplicationDbContext _context;
        private ValidateToken validator;

        public CustomUserController()
        {
            _context = new ApplicationDbContext();
            validator = new ValidateToken();
        }

        // GET api/user
        [Route("api/user/")]
        public IHttpActionResult Get()
        {
            var userlist = _context.CustomUsers.Include(u => u.Rol).ToList().Select(x=> new{ x.Id ,x.UserName , Rol= x.Rol.Name });
            return Ok(userlist);

        }

        // GET api/user/5
        [Route("api/user/{id}")]
        public IHttpActionResult Get(int id)
        {
            CustomUser userInDB = null;

            userInDB = _context.CustomUsers.Include(u => u.Rol).FirstOrDefault(d => d.Id == id);

            if (userInDB == null)
                return NotFound();
            dynamic respose = new JObject();
            respose.Id = userInDB.Id;
            respose.UserName = userInDB.UserName;
            respose.RolId = userInDB.RolId;

            return Ok(respose);
        }

        // POST: /api/user/
        [HttpPost]
        [Route("api/user/")]
        public IHttpActionResult Register([FromBody]CustomUser user)
        {
            if (!ModelState.IsValid)
                return BadRequest();

            PasswordHash hash = new PasswordHash(user.Password);
            byte[] hashBytes = hash.ToArray();
            user.Id = _context.Database.SqlQuery<int>("SELECT \"rrhh_User_sqs\".nextval FROM DUMMY;").ToList()[0];
            //user.Password = Encoding.ASCII.GetString(hashBytes);
            user.Password = Convert.ToBase64String(hashBytes);
            user.Token = validator.getToken(user);
            user.active = true;
            user.TokenCreatedAt = DateTime.Now;
            user.RefreshToken = validator.getRefreshToken(user);
            user.RefreshTokenCreatedAt = DateTime.Now;
            _context.CustomUsers.Add(user);
            _context.SaveChanges();

            dynamic respose = new JObject();
            respose.Id = user.Id;
            respose.UserName = user.UserName;
            respose.Token = user.Token;
            respose.RefreshToken = user.RefreshToken;
            respose.active = user.active;

            return Created(new Uri(Request.RequestUri + "/" + respose.Id), respose);
        }



        // PUT api/People/5
        [HttpPut]
        [Route("api/user/{id}")]
        public IHttpActionResult Put(int id, [FromBody]CustomUser user)
        {
            if (!ModelState.IsValid)
                return BadRequest();

            CustomUser userInDB = _context.CustomUsers.FirstOrDefault(d => d.Id == id);
            if (userInDB == null)
                return NotFound();
            userInDB.UserName = user.UserName;
            userInDB.RolId = user.RolId;
            userInDB.active = user.active;
            _context.SaveChanges();

            dynamic respose = new JObject();
            respose.Id = userInDB.Id;
            respose.UserName = userInDB.UserName;

            return Ok(respose);
        }

        [HttpPost]
        [Route("api/user/ChangePassword/{id}")]
        //POST api/user/ChangePassword/5
        public IHttpActionResult ChangePassword(int id, [FromBody]JObject credentials)
        {
            if (credentials["oldpassword"] == null || credentials["newpassword"] == null || credentials["newpassword2"] == null)
                return BadRequest();

            string oldpassword = credentials["oldpassword"].ToString();
            string newpassword = credentials["newpassword"].ToString();
            string newpassword2 = credentials["newpassword2"].ToString();

            CustomUser user = _context.CustomUsers.FirstOrDefault(u => u.Id == id);
            if (user == null)
                return Unauthorized();
     
            if (newpassword != newpassword2)
                return BadRequest("Las Contraseñas no coinciden");
            

            PasswordHash hash = new PasswordHash(newpassword);
            byte[] hashBytes = hash.ToArray();

            //string hashnewpassword = System.Text.Encoding.Default.GetString(hashBytes, 0, hashBytes.Length); ;
            string hashnewpassword = Convert.ToBase64String(hashBytes);

            hashBytes = Encoding.Default.GetBytes(user.Password);
            hashBytes = Convert.FromBase64String(user.Password);
            hash = new PasswordHash(hashBytes);

            if (hash.Verify(newpassword))
                return BadRequest("La nueva Contraseña no pude ser igual a la Contraseña actual");

           // if (!hash.Verify(oldpassword))
            //    return Unauthorized();

            user.Password = hashnewpassword;
            user.Token = validator.getToken(user);
            user.TokenCreatedAt = DateTime.Now;
            user.RefreshToken = validator.getRefreshToken(user);
            user.RefreshTokenCreatedAt = DateTime.Now;

            _context.SaveChanges();

            dynamic respose = new JObject();
            respose.Id = user.Id;
            respose.UserName = user.UserName;
            respose.Token = user.Token;
            respose.RefreshToken = user.RefreshToken;

            return Ok(respose);
        }

        // DELETE api/user/5
        [HttpPost]
        [Route("api/user/ChangeStatus")]
        public IHttpActionResult ChangeStatus(int id)
        {
            var userInDB = _context.CustomUsers.FirstOrDefault(d => d.Id == id);
            if (userInDB == null)
                return NotFound();

            _context.CustomUsers.Remove(userInDB);
            _context.SaveChanges();
            return Ok();
        }
    }
}
