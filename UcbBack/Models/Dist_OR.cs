using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace UcbBack.Models
{
    public class Dist_OR
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int Id { set; get; }
        public string Document { set; get; }
        public string FullName { set; get; }
        public string Segment { set; get; }
        public decimal HaberMensual { set; get; }
        public decimal OtrosIngresos { set; get; }
        public decimal TotalGanado { set; get; }

    }
}