using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Xml.Serialization;
using System.Data.Entity;
using System.Security.Cryptography;
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
                    .Include(x => x.Dependency.OrganizationalUnit)
                    .Include(x => x.Link)
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
                    .Include(x => x.Dependency.OrganizationalUnit)
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
            if (manager == null)
                return null;
            if (manager.CUNI == this.CUNI && contract.Dependency.ParentId != 0)
            {
                manager = _contex.ContractDetails
                    .Include(x => x.Positions)
                    .Include(x => x.People)
                    .Where(x => x.DependencyId == contract.Dependency.ParentId
                                && x.StartDate <= date
                                && (x.EndDate == null || x.EndDate >= date)).OrderBy(x => x.Positions.LevelId).Select(x=>x.People).FirstOrDefault();
            }

            if (manager == null)
                manager = null;
            return manager;
        }

        public People GetLastManagerAuthorizator(ApplicationDbContext _contex = null, DateTime? date = null)
        {
            _contex = _contex == null ? new ApplicationDbContext() : _contex;
            People manager;
            var contract = GetLastContract(_contex, date);
            date = date == null ? DateTime.Now : date;
            manager = _contex.ContractDetails
                .Include(x => x.Positions)
                .Include(x => x.Positions.Level)
                .Include(x => x.People)
                .Where(x => x.DependencyId == contract.DependencyId
                            && x.StartDate <= date
                            && (x.EndDate == null || x.EndDate >= date)
                            && (
                                x.Positions.Level.Cod == "N1" || // RECTOR NACIONAL
                                x.Positions.Level.Cod == "N2" || // VICERRECTOR NACIONAL
                                x.Positions.Level.Cod == "N3" || // SECRETARIO GENERAL NACIONAL
                                x.Positions.Level.Cod == "N4" || // AUDITOR INTERNO NACIONAL
                                x.Positions.Level.Cod == "N5" || // DIRECTOR NACIONAL
                                x.Positions.Level.Cod == "N6" || // COORDINADOR NACIONAL
                                x.Positions.Level.Cod == "R1" || // RECTOR REGIONAL
                                x.Positions.Level.Cod == "R3" || // SECRETARIA ACADEMICA || DIRECTOR REGIONAL || DIRECTOR AREA
                                x.Positions.Level.Cod == "R4" || // JEFE DEPARTAMENTO
                                x.Positions.Level.Cod == "R5" || // JEFE UNIDAD
                                x.Positions.Level.Cod == "G1" || // DECANO
                                x.Positions.Level.Cod == "G2" || // DIRECTOR
                                x.Positions.Level.Cod == "G3" || // JEFE ACADEMICO
                                x.Positions.Level.Cod == "G4"    // COORDINADOR ACADEMICO
                                )
                            ).OrderBy(x => x.Positions.LevelId).ThenBy(x=>x.StartDate).Select(x => x.People).FirstOrDefault();
            // case when manager of area
            var parentId = contract.Dependency.ParentId;
            while ((manager == null || manager.CUNI == this.CUNI) && parentId != 0)
            {
                manager = _contex.ContractDetails
                    .Include(x => x.Positions)
                    .Include(x => x.People)
                    .Where(x => x.DependencyId == parentId
                                && x.StartDate <= date
                                && (x.EndDate == null || x.EndDate >= date)
                                && (
                                    x.Positions.Level.Cod == "N1" ||
                                    x.Positions.Level.Cod == "N2" ||
                                    x.Positions.Level.Cod == "N3" ||
                                    x.Positions.Level.Cod == "N4" ||
                                    x.Positions.Level.Cod == "N5" ||
                                    x.Positions.Level.Cod == "N6" ||
                                    x.Positions.Level.Cod == "R1" ||
                                    x.Positions.Level.Cod == "R3" ||
                                    x.Positions.Level.Cod == "R4" ||
                                    x.Positions.Level.Cod == "R5" ||
                                    x.Positions.Level.Cod == "G1" ||
                                    x.Positions.Level.Cod == "G2" ||
                                    x.Positions.Level.Cod == "G3" ||
                                    x.Positions.Level.Cod == "G4"
                                )
                                ).OrderBy(x => x.Positions.LevelId).ThenBy(x => x.StartDate).Select(x => x.People).FirstOrDefault();
                parentId = _contex.Dependencies.FirstOrDefault(x => x.Id == parentId).ParentId;
            }

            if (manager == null)
                manager = null;
            return manager;
        }

        public static int GetNextId(ApplicationDbContext _context)
        {
            return _context.Database.SqlQuery<int>("SELECT \"" + CustomSchema.Schema + "\".\"rrhh_People_sqs\".nextval FROM DUMMY;").ToList()[0];
        }

        public void GetPerfilRendiciones(out int id,out string nom)
        {
            var reg = this.GetLastContract().Dependency.BranchesId;
            switch (reg)
            {
                case 2: //TJA
                    id = 4;
                    nom = "RENDICIONES_TJA-Moneda Local(BS)";
                    break;
                case 3: //CBB
                    id = 2;
                    nom = "RENDICIONES_CBB-Moneda Local(BS)";
                    break;
                case 7: //UCE
                    id = 7;
                    nom = "RENDICIONES_UCE-Moneda Local(BS)";
                    break;
                case 16: //SC
                    id = 3;
                    nom = "RENDICIONES_SCZ-Moneda Local(BS)";
                    break;
                case 17: //LP
                    id = 1;
                    nom = "RENDICIONES_LPZ-Moneda Local(BS)";
                    break;
                case 18: //EPC
                    id = 5;
                    nom = "RENDICIONES_EPC-Moneda Local(BS)";
                    break;
                case 22: //TEO
                    id = 6;
                    nom = "RENDICIONES_TEO-Moneda Local(BS)";
                    break;
                default:
                    id = 0;
                    nom = "";
                    break;
            }
        }

        public string GetContador()
        {
            var reg = this.GetLastContract().Dependency.BranchesId;
            switch (reg)
            {
                case 2: //TJA
                    return "DELGADILLO APARICIO EDGAR";
                case 3: //CBB
                    return "PEREDO GUMUCIO JONNY HAAMET";
                case 6: //UCE
                    return "AGUIRRE RIOS GLORIA DORIS";
                case 16: //SC
                    return "CAMACHO MORENO MARGARITA MARCIA";
                case 17: //LP
                    return "ALDUNATE MORALES NANCY JAEL";
                case 18: //EPC
                    return "ALIAGA CALCINA LIZ MARGOTH";
                case 22: //TEO
                    return "PEREDO GUMUCIO JONNY HAAMET";
                default:
                    return null;
            }
        }

        public void CreateInRendiciones(ApplicationDbContext _context)
        {
            
            _context = _context == null ? new ApplicationDbContext() : _context;
            var user = _context.CustomUsers.FirstOrDefault(x => x.PeopleId == this.Id);
            //var HashPass = user.AutoGenPass;
            var HashPass =
                _context.Database.SqlQuery<string>("select  to_varchar(hash_sha256(to_binary('" + user.AutoGenPass +
                                                   "'))) from dummy").ToList()[0].ToLower();

            int nextId = _context.Database.SqlQuery<int>("SELECT \"" + ConfigurationManager.AppSettings["RendicionesSchema"] + "\".\"DIMRENDUSUARIO_SEQ\".nextval FROM DUMMY;").ToList()[0];
            string query = "insert into  " +
                           " 	" + ConfigurationManager.AppSettings["RendicionesSchema"] + ".rend_u ( " +
                           " 			\"U_IdU\", " +
                           " 			\"U_Login\", " +
                           " 			\"U_Pass\", " +
                           " 			\"U_SuperUser\", " +
                           " 			\"U_AppRend\", " +
                           " 			\"U_AppExtB\", " +
                           " 			\"U_AppUpLA\", " +
                           " 			\"U_GenDocPre\", " +
                           " 			\"U_NomUser\", " +
                           " 			\"U_NomSup\", " +
                           " 			\"U_Estado\", " +
                           " 			\"U_AppConf\", " +
                           " 			\"U_CardCode\", " +
                           " 			\"U_CardName\" " +
                           " 		) " +
                           " 	values ( " +
                           " 			" + nextId + ", " +
                           " 			'" + user.UserPrincipalName + "', " +
                           " 			'" + HashPass + "', " +
                           " 			0, " +
                           " 			1, " +
                           " 			0, " +
                           " 			0, " +
                           " 			0, " +
                           " 			'" + this.GetFullName() + "', " +
                           " 			'" + this.GetContador() + "', " +
                           " 			1, " +
                           " 			0, " +
                           " 			'R" + this.CUNI + "', " +
                           " 			'R" + this.CUNI + "-" + this.GetFullName() + "' " +
                           " 		) ";

            var res = _context.Database.ExecuteSqlCommand(query);

            int idperfil;
            string nomperfil;
            this.GetPerfilRendiciones(out idperfil, out nomperfil);

            string query2 = "insert into \"" + ConfigurationManager.AppSettings["RendicionesSchema"] + "\".\"REND_PRM\" (U_IDUSUARIO, U_IDPERFIL, U_NOMBREPERFIL) " +
                            " values (" +
                            nextId+ ", " +
                            idperfil + ",'" +
                            nomperfil + "'" +
                            ")";

            var res2 = _context.Database.ExecuteSqlCommand(query2);

        }
    }
}