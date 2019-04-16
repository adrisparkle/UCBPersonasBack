using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using UcbBack.Logic;
using UcbBack.Models;
using UcbBack.Models.Auth;
using UcbBack.Models.Not_Mapped.CustomDataAnnotations;
using UcbBack.Models.ViewMoldes;

namespace UcbBack.Controllers
{
    public class CivilController : ApiController
    {

        private ApplicationDbContext _context;
        private ValidateAuth auth;


        public CivilController()
        {
            _context = new ApplicationDbContext();
            auth = new ValidateAuth();
        }

        // GET api/Level
        public IHttpActionResult Get()
        {
            // we get the Branches from SAP
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
                        " on c.\"SAPId\" = ocrd.\"CardCode\";";
            var rawresult = _context.Database.SqlQuery<Civil>(query).ToList();

            var user = auth.getUser(Request);

            var res = auth.filerByRegional(rawresult.AsQueryable(), user);

            return Ok(res);
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

        [HttpGet]
        [Route("api/Civil/findInSAP/{id}")]
        public IHttpActionResult findInSAP(int CardCode)
        {
            var user = auth.getUser(Request);
            var BP = findBPInSAP(CardCode, user);

            if (BP == null)
                return NotFound();

            if (BP == -1)
                return Unauthorized();

            return Ok(BP);
        }

        [NonAction]
        private dynamic findBPInSAP(int CardCode,CustomUser user)
        {
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
                        " and ocrd.\"CardType\" = 'S'" +
                        " where c.\"Id\"= " + CardCode + ";";
            var rawresult = _context.Database.SqlQuery<Civil>(query).ToList();

            if (rawresult.Count() == 0)
                return null;

            var res = auth.filerByRegional(rawresult.AsQueryable(), user);

            if (res.Count() == 0)
                return -1;

            return res.FirstOrDefault();
        }

        // POST api/Level
        [HttpPost]
        public IHttpActionResult Post([FromBody]Civil civil)
        {
            var user = auth.getUser(Request);
            var BP = findBPInSAP(Int32.Parse(civil.SAPId), user);

            if (!ModelState.IsValid)
                return BadRequest();
            if (BP == null)
                return NotFound();
            if (BP == -1)
                return Unauthorized();

            civil.Id = Civil.GetNextId(_context);
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
