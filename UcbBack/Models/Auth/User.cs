using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace UcbBack.Models.Auth
{
    public class User
    {
        public int Id { get; set; }
        public string UserName { get; set; }
        public string Password { get; set; }
        public Rol Rol { get; set; }
        public int RolId { get; set; }
    }
}