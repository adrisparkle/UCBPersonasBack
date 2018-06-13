using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace UcbBack.Models.Auth
{
    [Table("ADMNALRRHH.Access")]
    public class Access
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int Id { get; set; }
        public string Method { get; set; }
        public string Path { get; set; }
        public string Description { get; set; }
        public bool Public { get; set; }
    }
}