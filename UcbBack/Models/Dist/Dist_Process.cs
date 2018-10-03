using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace UcbBack.Models.Dist
{
    [Table("ADMNALRRHH.Dist_Process")]
    public class Dist_Process
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public long Id { set; get; }

        public DateTime UploadedDate { get; set; }

        public Branches Branches { get; set; }
        public int BranchesId { get; set; }

        public string mes { get; set; }
        public string gestion { get; set; }
        public string State { get; set; }
        public DateTime? RegisterDate { get; set; }
    }
}