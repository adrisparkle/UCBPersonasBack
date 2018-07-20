using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace UcbBack.Models
{
    [Table("ADMNALRRHH.Dist_Academic")]
    public class Dist_Academic
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int Id { set; get; }
        public string Document { get; set; }
        public string FirstName { get; set; }
        public string FirstSurName { get; set; }
        public string SecondSurName { get; set; }
        public string MariedSurName { get; set; }
        public string EmployeeType { get; set; }
        public string Periodo { get; set; }
        public string Sigla { get; set; }
        public string Paralelo { get; set; }
        public decimal AcademicHoursWeek { get; set; }
        public decimal AcademicHoursMonth { get; set; }
        public string IdentificadorPago { get; set; }
        public string CategoriaDocente { get; set; }
        public decimal CostoHora { get; set; }
        public decimal CostoMes { get; set; }
        public string CUNI { get; set; }
        public string Dependency { get; set; }
        public string PEI { get; set; }
        public string SAPParaleloUnit { get; set; }

        public decimal Porcentaje { get; set; }
        public int Matched { get; set; }
        public string segmentoOrigen { get; set; }
        [StringLength(2)]
        public string mes { get; set; }
        [StringLength(4)]
        public string gestion { get; set; }

        public Dist_File DistFile { get; set; }
        public int DistFileId { get; set; }
    }
}