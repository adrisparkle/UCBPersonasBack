using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.AccessControl;
using System.Web;

namespace UcbBack.Models
{
    public class Positions
    {
        public int Id { get; set; }
        public String NAME { get; set; }

        public int LevelId { get; set; }
        public Level Level { get; set; }

        public int BranchesId { get; set; }
        public Branches Branches { get; set; }
        
    }
}