using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace UcbBack.Models
{
    [Table("ADMNALRRHH.OrganizationalUnit")]
    public class OrganizationalUnit
    {
        public int Id { get; set; }
        public string Cod { get; set; }
        public string Name { get; set; }
    }
}