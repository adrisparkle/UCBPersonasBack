using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.Entity.Migrations;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web.Http;
using ClosedXML.Excel;
using Newtonsoft.Json.Linq;
using SAPbouiCOM;
using UcbBack.Logic;
using UcbBack.Logic.B1;
using UcbBack.Models;
using UcbBack.Models.Auth;
using UcbBack.Models.Not_Mapped;
using UcbBack.Models.Not_Mapped.CustomDataAnnotations;

namespace UcbBack.Controllers
{
    public class CivilController : ApiController
    {

        private ApplicationDbContext _context;
        private ValidateAuth auth;
        private ADClass AD;


        public CivilController()
        {
            _context = new ApplicationDbContext();
            auth = new ValidateAuth();
            AD = new ADClass();
        }

        // GET api/Level
        [HttpGet]
        [Route("api/CivilbyBranch/{id}")]
        public IHttpActionResult CivilbyBranch(int id)
        {
            var B1 = B1Connection.Instance();
            if (id != 0)
            {
                // we get the Branches from SAP
                var query = "select c.\"Id\", c.\"FullName\",c.\"SAPId\",c.\"NIT\",c.\"Document\",c.\"CreatedBy\",ocrd.\"BranchesId\" " +
                            "from " + CustomSchema.Schema + ".\"Civil\" c" +
                            " inner join " +
                            " (select ocrd.\"CardCode\", br.\"Id\" \"BranchesId\"" +
                            " from " + ConfigurationManager.AppSettings["B1CompanyDB"] + ".ocrd" +
                            " inner join " + ConfigurationManager.AppSettings["B1CompanyDB"] + ".crd8" +
                            " on ocrd.\"CardCode\" = crd8.\"CardCode\"" +
                            " inner join " + CustomSchema.Schema + ".\"Branches\" br" +
                            " on br.\"CodigoSAP\" = crd8.\"BPLId\"" +
                            " where ocrd.\"validFor\" = \'Y\'" +
                            " and crd8.\"DisabledBP\" = \'N\') ocrd" +
                            " on c.\"SAPId\" = ocrd.\"CardCode\"" +
                            " where ocrd.\"BranchesId\"=" + id + ";";
                var rawresult = _context.Database.SqlQuery<Civil>(query);

                var user = auth.getUser(Request);

                var res = auth.filerByRegional(rawresult.AsQueryable(), user);

                return Ok(res);
            }
            else
            {
                var user = auth.getUser(Request);
                var brs = AD.getUserBranches(user);
                var brsIds = brs.Select(x => x.Id);
                string StrIds = "";
                int n = brsIds.Count();
                int i = 0;
                foreach (var brid in brsIds)
                {
                    i++;
                    StrIds += brid + "" + (i==n?"":", ");
                    
                }


                var query = "select c.\"Id\", c.\"FullName\",c.\"SAPId\",c.\"NIT\",c.\"Document\",c.\"CreatedBy\",ocrd.\"BranchesId\" " +
                            "from " + CustomSchema.Schema + ".\"Civil\" c" +
                            " inner join " +
                            " (select ocrd.\"CardCode\", br.\"Id\" \"BranchesId\"" +
                            " from " + ConfigurationManager.AppSettings["B1CompanyDB"] + ".ocrd" +
                            " inner join " + ConfigurationManager.AppSettings["B1CompanyDB"] + ".crd8" +
                            " on ocrd.\"CardCode\" = crd8.\"CardCode\"" +
                            " inner join " + CustomSchema.Schema + ".\"Branches\" br" +
                            " on br.\"CodigoSAP\" = crd8.\"BPLId\"" +
                            " where ocrd.\"validFor\" = \'Y\'" +
                            " and crd8.\"DisabledBP\" = \'N\') ocrd" +
                            " on c.\"SAPId\" = ocrd.\"CardCode\"" +
                            " where ocrd.\"BranchesId\" in (" + StrIds + ");";
                var rawresult = _context.Database.SqlQuery<Civil>(query);
                var res = auth.filerByRegional(rawresult.AsQueryable(), user);
                return Ok(res);
            }
        }

        // GET api/Level/5
        public IHttpActionResult Get(int id)
        {
            var user = auth.getUser(Request);
            var query = "select c.*,ocrd.\"BranchesId\" from " + CustomSchema.Schema + ".\"Civil\" c" +
                        " inner join " +
                        " (select ocrd.\"CardCode\", br.\"Id\" \"BranchesId\"" +
                        " from " + ConfigurationManager.AppSettings["B1CompanyDB"] + ".ocrd" +
                        " inner join " + ConfigurationManager.AppSettings["B1CompanyDB"] + ".crd8" +
                        " on ocrd.\"CardCode\" = crd8.\"CardCode\"" +
                        " inner join " + CustomSchema.Schema + ".\"Branches\" br" +
                        " on br.\"CodigoSAP\" = crd8.\"BPLId\"" +
                        " where ocrd.\"validFor\" = \'Y\'" +
                        " and crd8.\"DisabledBP\" = \'N\') ocrd" +
                        " on c.\"SAPId\" = ocrd.\"CardCode\"" +
                        " where c.\"Id\"= " + id + ";";
            var rawresult = _context.Database.SqlQuery<Civil>(query).ToList();

            if (rawresult.Count() == 0)
                return NotFound();

            var res = auth.filerByRegional(rawresult.AsQueryable(), user);

            if (res.Count() == 0)
                return Unauthorized();

            return Ok(res.FirstOrDefault());
        }

        [HttpPost]
        [Route("api/CivilfindInSAP/")]
        public IHttpActionResult findInSAP(JObject CardCode)
        {
            var user = auth.getUser(Request);
            var BP = Civil.findBPInSAP(CardCode["CardCode"].ToString(), user,_context);
            
            if (BP == null)
                return NotFound();

            return Ok(BP.FirstOrDefault());
        }

        // POST api/Level
        [HttpPost]
        public IHttpActionResult Post([FromBody]Civil civil)
        {
            var user = auth.getUser(Request);


            var BP = Civil.findBPInSAP(civil.SAPId, user,_context);

            if (!ModelState.IsValid)
                return BadRequest();

            //todo validate BranchesId here

            if (BP == null)
                return Unauthorized();
            var a =AD.getUserBranches(user).Select(x => x.Id);
            var b = BP.Select(x => x.BranchesId);

            if (!a.Intersect(b).Any())
            {
                return Unauthorized();
            }

            var exists = _context.Civils.FirstOrDefault(x => x.SAPId == civil.SAPId);
            if (exists != null)
                //return Ok("Este Socio de Negocios ya existe como Civil.");
                return Conflict();

            civil.Id = Civil.GetNextId(_context);
            civil.CreatedBy = user.Id;
            _context.Civils.Add(civil);
            _context.SaveChanges();

            return Created(new Uri(Request.RequestUri + "/" + civil.Id), civil);
        }

        // DELETE api/Level/5
        [HttpDelete]
        public IHttpActionResult Delete(int id)
        {
            var levelInDB = _context.Levels.FirstOrDefault(d => d.Id == id);
            if (levelInDB == null)
                return NotFound();
            _context.Levels.Remove(levelInDB);
            _context.SaveChanges();
            return Ok();
        }

    }
}
