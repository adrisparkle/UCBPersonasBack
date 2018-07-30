using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace UcbBack.Models
{

    [Table("ADMNALRRHH.Resource")]
    public class Resource
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int Id { set; get; }
        public string Name { get; set; }
        public string Path { get; set; }

        public Module Module { get; set; }
        public int ModuleId { get; set; }
    }
}