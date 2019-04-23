using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;
using UcbBack.Models.Not_Mapped.CustomDataAnnotations;

namespace UcbBack.Models.Serv
{
    [CustomSchema("Serv_Process")]
    public class ServProcess
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int Id { set; get; }

        public int BranchesId { get; set; }
        public string FileType { get; set; }
        public DateTime CreatedAt { get; set; }
        public DateTime InSAPAt { get; set; }
        public string State { get; set; }
        public string SAPId { get; set; }

        public struct Serv_FileState
        {
            public const string Started = "INICIADO";
            public const string PendingApproval = "ESPERANDO APROVACION";
            public const string INSAP = "IN SAP";
            public const string ERROR = "ERROR";
            public const string Canceled = "RECHAZADO";
        }

        public struct Serv_FileType
        {
            public const string Varios = "VARIOS";
            public const string Proyectos = "PROYECTOS";
            public const string Pregrado = "PREGRADO";
            public const string Reemplazo = "REEMPLAZO";
        }
    }
}