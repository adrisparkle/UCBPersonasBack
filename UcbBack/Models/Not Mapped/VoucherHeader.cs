using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Web;

namespace UcbBack.Models.Not_Mapped
{
    [NotMapped]
    public class VoucherHeader
    {
        public string ParentKey { get; set; }
        public string ReferenceDate { get; set; }
        public string Memo { get; set; }
        public string TaxDate { get; set; }
        public string Series { get; set; }
        public string DueDate { get; set; }
    }

}