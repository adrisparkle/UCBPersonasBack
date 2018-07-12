using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace UcbBack.Models
{
    [Table("ADMNALRRHH.Dist_TipoEmpleado")]
    public class TipoEmpleadoDist
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int Id { set; get; }
        public GrupoContable GrupoContable { get; set; }
        public int GrupoContableId { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }

    }
}