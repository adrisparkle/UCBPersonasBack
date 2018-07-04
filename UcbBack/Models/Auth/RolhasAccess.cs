﻿using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace UcbBack.Models.Auth
{
    [Table("ADMNALRRHH.RolhasAccess")]
    public class RolhasAccess
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int Id { get; set; }
        public int Accessid { get; set; }
        public Access Access { get; set; }
        public int Rolid { get; set; }
        public Rol Rol { get; set; }
    }
}