using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace UcbBack.Models
{
    [Table("ADMNALRRHH.Dist_Process")]
    public class Dist_Process
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int Id { set; get; }

        public DateTime UploadedDate { get; set; }

        public Branches Branches { get; set; }
        public int BranchesId { get; set; }

        public string mes { get; set; }
        public string gestion { get; set; }
        public string State { get; set; }
    }
}