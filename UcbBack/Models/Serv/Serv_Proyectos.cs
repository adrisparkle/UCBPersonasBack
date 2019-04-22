using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;
using UcbBack.Models.Not_Mapped.CustomDataAnnotations;

namespace UcbBack.Models.Serv
{
    [CustomSchema("Serv_Proyectos")]
    public class Serv_Proyectos
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int Id { set; get; }

        public string CardCode { get; set; }
        public string CardName { get; set; }
        public int DependencyId { get; set; }
        public string PEI { get; set; }
        public string ProjectSAPCode { get; set; }
        public string ProjectSAPName { get; set; }
        public string Version { get; set; }
        public string Periodo { get; set; }
        public string JobType { get; set; }
        public string ServiceType { get; set; }
        public Decimal ContractAmount { get; set; }
        public Decimal IUE { get; set; }
        public Decimal IT { get; set; }
        public Decimal TotalAmount { get; set; }
        public string Comments { get; set; }

    }
}