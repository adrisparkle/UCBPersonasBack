using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;
using System.Xml.Serialization;

namespace UcbBack.Models
{
    [Table("ADMNALRRHH.People")]
    public class People
    {
        public int Id { get; set; }

        [MaxLength(10)]
        public string CUNI { get; set; }

        [MaxLength(15)]
        [Required]
        public string TypeDocument { get; set; }

        [MaxLength(15)]
        [Required]
        public string Document { get; set; }

        [MaxLength(5)]
        [Required]
        public string Ext { get; set; }

        [MaxLength(200)]
        [Required]
        public string Names { get; set; }

        [MaxLength(100)]
        [Required]
        public string FirstSurName { get; set; }

        [MaxLength(100)]
        [Required]
        public string SecondSurName { get; set; }

        [MaxLength(100)]
        public string MariedSurName { get; set; }

        [Column(TypeName = "date")]
        //[DisplayFormat(ApplyFormatInEditMode = true, DataFormatString = "{0:dd-MM-yyyy}")]
        [Required]
        public DateTime BirthDate { get; set; }

        [MaxLength(1)]
        [Required]
        public string Gender { get; set; }

        [MaxLength(20)]
        [Required]
        public string Nationality { get; set; }

        [MaxLength(250)]
        public string Photo { get; set; }

        [MaxLength(15)]
        public string PhoneNumber { get; set; }

        [MaxLength(30)]
        public string PersonalEmail { get; set; }

        [MaxLength(30)]
        public string UcbEmail { get; set; }

        [MaxLength(15)]
        public string OfficePhoneNumber { get; set; }

        [MaxLength(15)]
        public string OfficePhoneNumberExt { get; set; }

        [MaxLength(200)]
        public string HomeAddress { get; set; }

        [MaxLength(20)]
        public string AFP { get; set; }

        [MaxLength(30)]
        public string NUA { get; set; }

        [MaxLength(50)]
        public string Insurance { get; set; }

        [MaxLength(20)]
        public string InsuranceNumber { get; set; }

        public bool UseMariedSurName { get; set; }

        public bool UseSecondSurName { get; set; }
    }
}