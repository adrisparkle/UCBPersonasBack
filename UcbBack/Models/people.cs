using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;
using System.Xml.Serialization;
using System.Data.Entity;
using Microsoft.Ajax.Utilities;
using UcbBack.Models.Not_Mapped.CustomDataAnnotations;

namespace UcbBack.Models
{
    [CustomSchema("People")]
    public class People
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int Id { get; set; }

        [MaxLength(10, ErrorMessage = "Cadena de texto muy grande")]
        public string CUNI { get; set; }

        [MaxLength(15, ErrorMessage = "Cadena de texto muy grande")]
        //[Required]
        public string TypeDocument { get; set; }

        [MaxLength(15, ErrorMessage = "Cadena de texto muy grande")]
        [Required]
        public string Document { get; set; }

        [MaxLength(5, ErrorMessage = "Cadena de texto muy grande")]
        //[Required]
        public string Ext { get; set; }

        [MaxLength(200, ErrorMessage = "Cadena de texto muy grande")]
        [Required]
        public string Names { get; set; }

        [MaxLength(100, ErrorMessage = "Cadena de texto muy grande")]
        [Required]
        public string FirstSurName { get; set; }

        [MaxLength(100, ErrorMessage = "Cadena de texto muy grande")]
        public string SecondSurName { get; set; }

        [MaxLength(100, ErrorMessage = "Cadena de texto muy grande")]
        public string MariedSurName { get; set; }

        [Column(TypeName = "date")]
        //[DisplayFormat(ApplyFormatInEditMode = true, DataFormatString = "{0:dd-MM-yyyy}")]
        [Required]
        public DateTime BirthDate { get; set; }

        [MaxLength(1, ErrorMessage = "Cadena de texto muy grande")]
        [Required]
        public string Gender { get; set; }

        [MaxLength(20, ErrorMessage = "Cadena de texto muy grande")]
        //[Required]
        public string Nationality { get; set; }

        [MaxLength(250, ErrorMessage = "Cadena de texto muy grande")]
        public string Photo { get; set; }

        [MaxLength(15, ErrorMessage = "Cadena de texto muy grande")]
        public string PhoneNumber { get; set; }

        [MaxLength(30, ErrorMessage = "Cadena de texto muy grande")]
        public string PersonalEmail { get; set; }

        [MaxLength(30, ErrorMessage = "Cadena de texto muy grande")]
        public string UcbEmail { get; set; }

        [MaxLength(15, ErrorMessage = "Cadena de texto muy grande")]
        public string OfficePhoneNumber { get; set; }

        [MaxLength(15, ErrorMessage = "Cadena de texto muy grande")]
        public string OfficePhoneNumberExt { get; set; }

        [MaxLength(200, ErrorMessage = "Cadena de texto muy grande")]
        public string HomeAddress { get; set; }

        [MaxLength(20, ErrorMessage = "Cadena de texto muy grande")]
        public string AFP { get; set; }

        [MaxLength(30, ErrorMessage = "Cadena de texto muy grande")]
        public string NUA { get; set; }

        [MaxLength(50, ErrorMessage = "Cadena de texto muy grande")]
        public string Insurance { get; set; }

        [MaxLength(20, ErrorMessage = "Cadena de texto muy grande")]
        public string InsuranceNumber { get; set; }

        [MaxLength(250, ErrorMessage = "Cadena de texto muy grande")]
        public string DocPath { get; set; }

        public int? SAPCodeRRHH { get; set; }

        public bool UseMariedSurName { get; set; }

        public bool UseSecondSurName { get; set; }

        public bool Pending { get; set; }


        public string GetFullName()
        {
            return this.FirstSurName + " " +
                   (this.UseSecondSurName ? (this.SecondSurName + " ") : "") +
                   (this.UseMariedSurName ? (this.MariedSurName + " ") : "") +
                   this.Names;
        }

        public ContractDetail GetLastContract(ApplicationDbContext _contex=null, DateTime? date=null)
        {
            _contex = _contex == null ? new ApplicationDbContext() : _contex;
            ContractDetail contract;
            if (date != null)
            {
                contract = _contex.ContractDetails
                    .Include(x => x.Branches)
                    .Include(x => x.Positions)
                    .Include(x => x.Dependency)
                    .Include(x => x.Link)
                    //.Include(x => x.Dependency.OrganizationalUnitId)
                    .Include(x=>x.People)
                    .Where(x => x.CUNI == this.CUNI
                                         && x.StartDate <= date
                                         && (x.EndDate == null || x.EndDate.Value >= date)).OrderBy(x=>x.Positions.LevelId).FirstOrDefault();
            }
            else
            {
                contract = _contex.ContractDetails
                    .Include(x => x.Branches)
                    .Include(x => x.Positions)
                    .Include(x => x.Dependency)
                    .Include(x => x.Link)
                    //.Include(x => x.Dependency.OrganizationalUnitId)
                    .Include(x=>x.People)
                    .Where(x => x.CUNI == this.CUNI)
                    .OrderByDescending(x => x.EndDate == null ? 1 : 0).ThenByDescending(x => x.EndDate).ThenBy(x=>x.Positions.LevelId).FirstOrDefault();
            }
            return contract;
        }

        public People GetLastManager(ApplicationDbContext _contex = null, DateTime? date = null)
        {
            _contex = _contex == null ? new ApplicationDbContext() : _contex;
            People manager;
            var contract = GetLastContract(_contex, date);
            date = date == null ? DateTime.Now : date;
                manager = _contex.ContractDetails
                    .Include(x => x.Positions)
                    .Include(x => x.People)
                    .Where(x => x.DependencyId == contract.DependencyId 
                                && x.StartDate <= date
                                && (x.EndDate == null || x.EndDate >= date)).OrderBy(x => x.Positions.LevelId).Select(x=>x.People).FirstOrDefault();
            // case when manager of area
            if (manager.CUNI == this.CUNI && contract.Dependency.ParentId != 0)
            {
                manager = _contex.ContractDetails
                    .Include(x => x.Positions)
                    .Include(x => x.People)
                    .Where(x => x.DependencyId == contract.Dependency.ParentId
                                && x.StartDate <= date
                                && (x.EndDate == null || x.EndDate >= date)).OrderBy(x => x.Positions.LevelId).Select(x=>x.People).FirstOrDefault();
            }

            return manager;
        }

        public static int GetNextId(ApplicationDbContext _context)
        {
            return _context.Database.SqlQuery<int>("SELECT \"" + CustomSchema.Schema + "\".\"rrhh_People_sqs\".nextval FROM DUMMY;").ToList()[0];
        }
    }
}