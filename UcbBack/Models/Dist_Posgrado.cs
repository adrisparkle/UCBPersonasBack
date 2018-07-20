using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace UcbBack.Models
{
    [Table("ADMNALRRHH.Dist_Posgrado")]
    public class Dist_Posgrado
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int Id { set; get; }
        public string Document { set; get; }
        public string FirstName { get; set; }
        public string FirstSurName { get; set; }
        public string SecondSurName { get; set; }
        public string MariedSurName { get; set; }
        public string ProjectName { set; get; }
        public int Vesion { set; get; }
        public decimal TotalPagado { set; get; }
        public string Dependency { set; get; }
        public string CUNI { set; get; }
        public string ProjectType { set; get; }
        public string TipoTarea { set; get; }
        public string PEI { set; get; }
        public string Periodo { set; get; }
        public string ProjectCode { set; get; }
        public decimal Porcentaje { get; set; }

        public string segmentoOrigen { get; set; }
        [StringLength(2)]
        public string mes { get; set; }
        [StringLength(4)]
        public string gestion { get; set; }

        public Dist_File DistFile { get; set; }
        public int DistFileId { get; set; }
    }
}