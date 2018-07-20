using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;
using UcbBack.Models.Auth;

namespace UcbBack.Models
{
    [Table("ADMNALRRHH.Dist_File")]
    public class Dist_File
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int Id { set; get; }

        public DateTime UploadedDate { get; set; }
        public string Name { get; set; }

        public Dist_FileType DistFileType { get; set; }
        public int DistFileTypeId { get; set; }

        public string State { get; set; }

        public Dist_Process DistProcess { get; set; }
        public int DistProcessId { get; set; }

        public CustomUser CustomUser { get; set; }
        public int CustomUserId { get; set; }

    }
}