using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;
using UcbBack.Models.Not_Mapped.CustomDataAnnotations;

namespace UcbBack.Models
{
    public class Chart
    {
        public string CodigoPadre { set; get; }
        public string NombrePadre { set; get; }
        public string CodigoDep { set; get; }
        public string NombreDep { set; get; }
        public string CodigoUO { set; get; }
        public string NombreUO { set; get; }
        public string Regional { get; set; }
    }
}