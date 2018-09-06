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

        [MaxLength(20, ErrorMessage = "Cadena de texto muy grande")]
        [Required]
        public string Name { get; set; }

        [MaxLength(10, ErrorMessage = "Cadena de texto muy grande")]
        [Required]
        public string Abr { get; set; }

        [Required]
        [MaxLength(100, ErrorMessage = "Cadena de texto muy grande")]
        public string ADGroupName { get; set; }

        [Required]
        [MaxLength(100, ErrorMessage = "Cadena de texto muy grande")]
        public string ADOUName { get; set; }

        [Required]
        [MaxLength(100, ErrorMessage = "Cadena de texto muy grande")]
        public string SerieComprobanteContalbeSAP { get; set; }
        
        [Required]
        [MaxLength(100, ErrorMessage = "Cadena de texto muy grande")]
        public string InitialsInterRegional { get; set; }

        [Required]
        [MaxLength(100, ErrorMessage = "Cadena de texto muy grande")]
        public string SocioGenericDerechosLaborales { get; set; }

        public Dependency Dependency { get; set; }

        public int? DependencyId { get; set; }

        public string CodigoSAP { get; set; }

    }
}