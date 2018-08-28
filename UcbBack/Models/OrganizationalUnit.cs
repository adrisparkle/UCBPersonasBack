using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace UcbBack.Models
{
    [Table("ADMNALRRHH.OrganizationalUnit")]
    public class OrganizationalUnit
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int Id { set; get; }

        [MaxLength(10, ErrorMessage = "Cadena de texto muy grande")]
        [Required]
        public string Cod { get; set; }

        [MaxLength(150, ErrorMessage = "Cadena de texto muy grande")]
        [Required]
        public string Name { get; set; }
    }
}