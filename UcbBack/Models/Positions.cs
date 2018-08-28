using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Security.AccessControl;
using System.Web;

namespace UcbBack.Models
{
    [Table("ADMNALRRHH.Position")]
    public class Positions
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int Id { set; get; }

        [MaxLength(50, ErrorMessage = "Cadena de texto muy grande")]
        [Required]
        public String Name { get; set; }

        [Required]
        public int LevelId { get; set; }
        public Level Level { get; set; }

        [Required]
        public int BranchesId { get; set; }
        public Branches Branches { get; set; }
        
    }
}