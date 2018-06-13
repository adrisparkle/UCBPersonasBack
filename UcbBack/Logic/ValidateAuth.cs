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
            else
            {
                return true;
            }
        }

        public bool hasAccess(int id, string path,string method)
        {
            CustomUser user = _context.CustomUsers.FirstOrDefault(u => u.Id == id);

            if (user == null)
            {
                return false;
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
            Access access = _context.Accesses.FirstOrDefault(a => a.Path == path && a.Method == method && a.Public==true);

            if (access == null)
            {
                return false;
            }

            return true;
        }

        public bool shallYouPass(int id, string token, string path, string method)
        {
            bool ispublic = isPublic(path, method);
            bool isauthenticated = isAuthenticated(id, token);
            bool hasaccess = hasAccess(id,path,method);
            return (ispublic || (isauthenticated&&hasaccess));
        }
    }
}