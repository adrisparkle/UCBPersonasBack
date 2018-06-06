using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Owin;
using Owin;

[assembly: OwinStartup(typeof(UcbBack.Startup))]

namespace UcbBack
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            app.Use(async (environment, next) =>
                {
                    var req = environment.Request;
                    string endpoint = environment.Request.Path.ToString();
                    string verb = environment.Request.Method;
                    if (endpoint == "/api/values/2" && verb == "GET")
                    {
                        environment.Response.StatusCode = 404;
                    }
                    else
                    await next();
                }
            );
            ConfigureAuth(app);
        }
    }
}
