using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace UcbBack.Models.ViewMoldes
{
    [NotMapped]
    public class UserViewModel
    {
        public int? Id { get; set; }
        public int? SAPCodeRRHH { get; set; }
        public string CUNI { get; set; }
        public string Document { get; set; }
        public string FullName { get; set; }
        public string TipoLicenciaSAP { get; set; }
        public bool? CajaChica { get; set; }
        public bool? SolicitanteCompras { get; set; }
        public bool? AutorizadorCompras { get; set; }
        public bool? Rendiciones { get; set; }
        public string UcbEmail { get; set; }
        public string UserPrincipalName { get; set; }

        public string Dependency { get; set; }
        public string DependencyCod { get; set; }

        public string OUCod { get; set; }
        public string OUName { get; set; }
        public string Positions { get; set; }
        public string Dedication { get; set; }
        public string Linkage { get; set; }

        public int? AuthSAPCodeRRHH { get; set; }
        public string AuthCUNI { get; set; }
        public string AuthFullName { get; set; }
        public string AuthPositions { get; set; }

        public string Branches { get; set; }
        public string AutoGenPass { get; set; }
    }
}