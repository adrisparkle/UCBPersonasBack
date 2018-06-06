using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;
using System.Xml.Serialization;

namespace UcbBack.Models
{
    public class People
    {
        public int Id { get; set; }

        [MaxLength(11)]
        public string COD_UCB { get; set; }

        [MaxLength(25)]
        [Required]
        public string TYPE_DOCUMENT { get; set; }

        [MaxLength(20)]
        [Required]
        public string DOCUMENTO { get; set; }

        [MaxLength(20)]
        [Required]
        public string ISSUED { get; set; }

        [MaxLength(100)]
        [Required]
        public string NAMES { get; set; }

        [MaxLength(50)]
        [Required]
        public string FIRSTSURNAME { get; set; }

        [MaxLength(50)]
        [Required]
        public string SECONDSURNAME { get; set; }

        [MaxLength(50)]
        public string MARIEDSURNAME { get; set; }

        [Column(TypeName = "date")]
        [Required]
        public DateTime? BIRTHDATE { get; set; }

        [MaxLength(1)]
        [Required]
        public string GENDER { get; set; }

        [MaxLength(20)]
        [Required]
        public string NATIONALITY { get; set; }

        [MaxLength(100)]
        public string PHOTO { get; set; }

        [MaxLength(25)]
        public string PHONENUMBER { get; set; }

        [MaxLength(50)]
        public string PERSONALEMAIL { get; set; }

        [MaxLength(50)]
        public string UCBMAIL { get; set; }

        [MaxLength(25)]
        public string OFFICEPHONENUMBER { get; set; }

        [MaxLength(200)]
        public string HOMEADDRESS { get; set; }

        [MaxLength(30)]
        public string MARITALSTATUS { get; set; }
    }
}