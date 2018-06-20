using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace UcbBack.Models
{
    [Table("ADMNALRRHH.Dependency")]
    public class Dependency
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int Id { set; get; }
        public string Cod { get; set; }
        public string Name { get; set; }
        public int Parent { get; set; }
        public int OrganizationalUnitId { get; set; }
        public OrganizationalUnit OrganizationalUnit { get; set; }
    }
}