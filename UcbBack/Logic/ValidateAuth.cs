using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using UcbBack.Models;
using UcbBack.Models.Auth;

namespace UcbBack.Logic
{
    public class ValidateAuth
    {
        private ApplicationDbContext _context;
        private ValidatePerson validator;
        public int tokenLife = 10;//15*60;
        public int refeshtokenLife = 60;//4*60*60;


        public ValidateAuth()
        {
            _context = new ApplicationDbContext();
        }

        public bool isAuthenticated(int id, string token)
        {
            CustomUser user = _context.CustomUsers.FirstOrDefault(u=> u.Id ==id);
            
            if (user == null || user.Token != token)
            {
                return false;
            }
            var now = DateTime.Now;
            if (user.TokenCreatedAt == null)
                return false;
            int seconds = (int)now.Subtract(user.TokenCreatedAt.Value).TotalSeconds;
            if (seconds > tokenLife)
            {
                user.Token = null;
                user.TokenCreatedAt = null;
                _context.SaveChanges();
                return false;
            }
            
            return true;
        }

        public bool hasAccess(int id, string path,string method)
        {
            CustomUser user = _context.CustomUsers.FirstOrDefault(u => u.Id == id);

            if (user == null)
            {
                return false;
            }

            if (user.RolId == 1)
            {
                return true;
            }

            Access access = _context.Accesses.FirstOrDefault(a => a.Path == path && a.Method == method);

            if (access == null)
            {
                return false;
            }

            RolhasAccess rolhasAccess =
                _context.RolshaAccesses.FirstOrDefault(ra => ra.Accessid == access.Id && ra.Rolid == user.RolId);
            if (rolhasAccess == null)
            {
                return false;
            }

            return true;
        }

        public bool isPublic(string path, string method)
        {
            Access access = _context.Accesses.FirstOrDefault(a =>
                string.Equals(a.Path.ToUpper(), path.ToUpper()) && a.Method == method);



            return access==null?false:access.Public;
        }

        public bool shallYouPass(int id, string token, string path, string method)
        {
            path = path.EndsWith("/")? path.Substring(0, path.Length - 1):path;

            bool ispublic = isPublic(path, method);
            bool isauthenticated = isAuthenticated(id, token);
            bool hasaccess = hasAccess(id,path,method);
            return (ispublic || (isauthenticated && hasaccess));
        }
    }
}