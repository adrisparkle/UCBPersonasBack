using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace UcbBack.Models
{
    [Table("ADMNALRRHH.ContractDetail")]
    public class ContractDetail
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int Id { set; get; }
        
        public string CUNI { get; set; }

        public int PeopleId { get; set; }
        public People People { get; set; }

        public int DependencyId { get; set; }
        public Dependency Dependency { get; set; }

        public int PositionsId { get; set; }
        public Positions Positions { get; set; }

        public string PositionDescription { get; set; }
        public string Dedication { get; set; }
        public string Linkage { get; set; }

        public int BranchesId { get; set; }
        public Branches Branches { get; set; }

        [Column(TypeName = "date")]
        public DateTime StartDate { get; set; }
        [Column(TypeName = "date")]
        public DateTime? EndDate { get; set; }
        public bool AI { get; set; }

        public string Cause { get; set; }
    }
}