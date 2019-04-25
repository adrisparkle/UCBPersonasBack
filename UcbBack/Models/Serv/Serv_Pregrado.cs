using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;
using UcbBack.Models.Not_Mapped.CustomDataAnnotations;

namespace UcbBack.Models.Serv
{
    [CustomSchema("Serv_Pregrado")]
    public class Serv_Pregrado
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int Id { set; get; }
        public int Serv_ProcessId { get; set; }

        public string CardCode { get; set; }
        public string CardName { get; set; }
        public int DependencyId { get; set; }
        public string PEI { get; set; }
        public string Memo { get; set; }
        public string Carrera { get; set; }
        public string DocumentNumber { get; set; }
        public string Student { get; set; }
        public string JobType { get; set; }
        public int Hours { get; set; }
        public Decimal CostPerHour { get; set; }
        public string ServiceType { get; set; }
        public Decimal ContractAmount { get; set; }
        public Decimal IUE { get; set; }
        public Decimal IT { get; set; }
        public Decimal TotalAmount { get; set; }
        public string Comments { get; set; }
    }
}