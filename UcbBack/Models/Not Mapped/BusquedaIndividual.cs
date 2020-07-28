using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Web;

namespace UcbBack.Models.Not_Mapped
{
    [NotMapped]
    public class BusquedaIndividual
    {
        public string Id { get; set; }
        public string PeopleId { get; set; }
        public string CUNI { get; set; }
        public string Document { get; set; }
        public string FullName { get; set; }
        public string Positions { get; set; }
        public string Linkage { get; set; }
        public string Dependency { get; set; }
        public string Branches { get; set; }
        public int BranchesId { get; set; }
        public string Status { get; set; }
    }
}