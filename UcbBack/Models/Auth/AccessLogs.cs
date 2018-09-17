using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace UcbBack.Models.Auth
{
    [Table("ADMNALRRHH.AccessLogs")]
    public class AccessLogs
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public long Id { get; set; }
        public string Method { get; set; }
        public string Path { get; set; }
        public int UserId { get; set; }
        public int AccessId { get; set; }
        public bool Success { get; set; }
        public string ResponseCode { get; set; }
    }
}