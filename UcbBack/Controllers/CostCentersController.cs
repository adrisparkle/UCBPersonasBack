using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using Newtonsoft.Json.Linq;
using UcbBack.Logic.B1;
using UcbBack.Models;

namespace UcbBack.Controllers
{
    public class CostCentersController : ApiController
    {
        private ApplicationDbContext _context;
        private B1Connection B1conn;
        public CostCentersController()
        {
            B1conn = B1Connection.Instance;
            _context = new ApplicationDbContext();
        }

        [HttpGet]
        [Route("api/CostCenters/OrganizationalUnits")]
        public IHttpActionResult OrganizationalUnits()
        {
            var y = B1conn.getCostCenter(B1Connection.Dimension.OrganizationalUnit,col:"*").Cast<JObject>();
            return Ok(y);
        }
        [HttpGet]
        [Route("api/CostCenters/PEI")]
        public IHttpActionResult PEI()
        {
            var y = B1conn.getCostCenter(B1Connection.Dimension.PEI, col: "*").Cast<JObject>();
            return Ok(y);
        }
        [HttpGet]
        [Route("api/CostCenters/PlanDeEstudios")]
        public IHttpActionResult PlanDeEstudios()
        {
            var y = B1conn.getCostCenter(B1Connection.Dimension.PlanAcademico, col: "*").Cast<JObject>();
            return Ok(y);
        }
        [HttpGet]
        [Route("api/CostCenters/Paralelo")]
        public IHttpActionResult Paralelo()
        {
            var y = B1conn.getCostCenter(B1Connection.Dimension.Paralelo, col: "*").Cast<JObject>();
            return Ok(y);
        }
        [HttpGet]
        [Route("api/CostCenters/Periodo")]
        public IHttpActionResult Periodo()
        {
            var y = B1conn.getCostCenter(B1Connection.Dimension.Periodo, col: "*").Cast<JObject>();
            return Ok(y);
        }
        [HttpGet]
        [Route("api/CostCenters/Proyectos")]
        public IHttpActionResult Proyectos()
        {
            var y = B1conn.getProjects("*");
            return Ok(y);
        }
    }
}
