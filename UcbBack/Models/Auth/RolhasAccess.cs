using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace UcbBack.Models.Auth
{
    public class RolhasAccess
    {
        public int Accessid { get; set; }
        public Access Access { get; set; }
        public int Rolid { get; set; }
        public Rol Rol { get; set; }
    }
}