using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace UcbBack.Models
{
    [Table("ADMNALRRHH.BRANCHES")]
    public class Branches
    {
        public int Id { get; set; }
        [MaxLength(20)]
        [Required]
        public string NAME { get; set; }
        [MaxLength(10)]
        [Required]
        public string ABR { get; set; }
    }
}