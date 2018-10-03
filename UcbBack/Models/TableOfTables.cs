using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;
using UcbBack.Models.Not_Mapped.CustomDataAnnotations;

namespace UcbBack.Models
{
    [CustomSchema("TableOfTables")]
    public class TableOfTables
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public Int64 Id { set; get; }
        public string Type { get; set; }
        public string Value { get; set; }

        
    }
    public struct TableOfTablesTypes
    {
        public static string Linkage = "VINCULACION";
        public static string Dedication = "DEDICACION";
        public static string CauseDesvinculation = "CAUSA DESVINCULACION";
        public static string AFP = "AFP";
        public static string ProcessState = "PROCESS STATE";
        public static string FileState = "FILE STATE";
    }
}