using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace UcbBack.Models.Auth
{
    [Table("ADMNALRRHH.Rol")]
    public class Rol
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int Id { get; set; }

        [Required]
        [MaxLength(50, ErrorMessage = "Cadena de texto muy grande")]
        public string Name { get; set; }

        [Required]
        public int Level { get; set; }

        [Required]
        [MaxLength(100, ErrorMessage = "Cadena de texto muy grande")]
        public string ADGroupName { get; set; }

        public Resource Resource { get; set; }
        public int ResourceId { get; set; }

        //public virtual ICollection<RolhasAccess> RolhasAccesses { get; set; }
    }
}