using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;
using UcbBack.Models.Not_Mapped.CustomDataAnnotations;

namespace UcbBack.Models.Auth
{
    [CustomSchema("User")]
    public class CustomUser
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int Id { get; set; }
        [Required]
        public string UserPrincipalName { get; set; }
        public People People { get; set; }
        public int PeopleId { get; set; }
        public string Token { get; set; }
        public string RefreshToken { get; set; }
        public DateTime ?TokenCreatedAt { get; set; }
        public DateTime ?RefreshTokenCreatedAt { get; set; }

        public static int GetNextId(ApplicationDbContext _context)
        {
            return _context.Database.SqlQuery<int>("SELECT \"" + CustomSchema.Schema + "\".\"rrhh_User_sqs\".nextval FROM DUMMY;").ToList()[0];
        }
    }
}