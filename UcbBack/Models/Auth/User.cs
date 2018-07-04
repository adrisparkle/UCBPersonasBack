using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace UcbBack.Models.Auth
{
    [Table("ADMNALRRHH.User")]
    public class CustomUser
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int Id { get; set; }
        [Required]
        public string UserName { get; set; }
        [MaxLength(50)]
        public string Password { get; set; }
        public Rol Rol { get; set; }
        public int RolId { get; set; }
        public string Token { get; set; }
        public string RefreshToken { get; set; }
        public DateTime? TokenCreatedAt { get; set; }
        public DateTime? RefreshTokenCreatedAt { get; set; }
        public Boolean active { get; set; }
    }
}