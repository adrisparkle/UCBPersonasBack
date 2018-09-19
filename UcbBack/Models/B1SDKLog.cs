using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace UcbBack.Models
{
    public class B1SDKLog
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int Id { set; get; }
        public string BusinessObject { get; set; }
        public int UserId { get; set; }
        public string ObjectId { get; set; }
        public string ErrorMessage { get; set; }
        public string ErrorCode { get; set; }
        public bool Success { get; set; }
    }
}