using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;
using UcbBack.Models.Not_Mapped.CustomDataAnnotations;

namespace UcbBack.Models
{
    [CustomSchema("Civil")]
    public class Civil
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int Id { get; set; }
        public string FullName { get; set; }
        public string SAPId { get; set; }
        public string NIT { get; set; }
        public string Document { get; set; }
        public int CreatedBy { get; set; }
        [NotMapped] public int BranchesId { get; set; }
        [NotMapped] public Branches Branches { get; set; }

        public static int GetNextId(ApplicationDbContext _context)
        {
            return _context.Database.SqlQuery<int>("SELECT \"" + CustomSchema.Schema + "\".\"rrhh_Civil_sqs\".nextval FROM DUMMY;").ToList()[0];
        }
    }
}