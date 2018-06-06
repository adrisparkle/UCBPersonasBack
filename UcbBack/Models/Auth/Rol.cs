using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace UcbBack.Models.Auth
{
    public class Rol
    {
        public int Id { get; set; }
        public int Name { get; set; }
        public int Nivel { get; set; }
        public virtual ICollection<RolhasAccess> RolhasAccesses { get; set; }
    }
}