using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace UcbBack.Models
{
    public class Dist_Pregrado
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int Id { set; get; }

        public string Document { get; set; }
        public string FullName { get; set; }
        public string Carrera { get; set; }
        public Decimal TotalBruto { get; set; }
        public Decimal Porcentaje { get; set; }
        public Decimal TotalNeto { get; set; }

    }
}