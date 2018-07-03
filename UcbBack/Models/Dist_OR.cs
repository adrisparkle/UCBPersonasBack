using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace UcbBack.Models
{
    [Table("ADMNALRRHH.Dist_OR")]
    public class Dist_OR
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int Id { set; get; }
        public string Document { set; get; }
        public string FullName { set; get; }
        public string segmento { set; get; }
        public decimal TotalGanado { set; get; }
        public string CUNI { set; get; }
        public string Dependency { set; get; }
        public string PEI { set; get; }
        public string PlanEstudios { set; get; }
        public string Paralelo { set; get; }
        public string Periodo { set; get; }
        public string Project { set; get; }

        public decimal Porcentaje { get; set; }
        public string segmentoOrigen { get; set; }
        public string mes { get; set; }
        public string gestion { get; set; }

    }
}