using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace UcbBack.Models
{
    [Table("ADMNALRRHH.Dist_Discounts")]
    public class Dist_Discounts
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int Id { set; get; }
        public string BussinesPartner { get; set; }
        public string Name { get; set; }
        public string Type { get; set; }
        public decimal Total { get; set; }
        public string segmentoOrigen { get; set; }
        public string mes { get; set; }
        public string gestion { get; set; }
    }
}