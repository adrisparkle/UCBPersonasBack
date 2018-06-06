using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace UcbBack.Models.Auth
{
    public class Access
    {
        public int Id { get; set; }
        public string Method { get; set; }
        public string Path { get; set; }
    }
}