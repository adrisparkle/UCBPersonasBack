using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace UcbBack.Models
{
    [Table("ADMNALRRHH.Branches")]
    public class Branches
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int Id { set; get; }
        [MaxLength(20)]
        [Required]
        public string Name { get; set; }
        [MaxLength(10)]
        [Required]
        public string Abr { get; set; }
    }
}