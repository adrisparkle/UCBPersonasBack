using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace UcbBack.Models
{
    [Table("ADMNALRRHH.Contract")]
    public class Contract
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int Id { set; get; }
        public string CUNI { get; set; }
        public int PeopleId { get; set; }
        public People People { get; set; }
        [Column(TypeName = "date")]
        public DateTime StartDate { get; set; }
        [Column(TypeName = "date")]
        public DateTime? EndDate { get; set; }

    }
}