using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web.WebPages.Scope;
using Microsoft.Owin;
using Owin;
using UcbBack.Logic;

[assembly: OwinStartup(typeof(UcbBack.Startup))]

namespace UcbBack
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            bool debugmode = true;
            app.Use(async (environment, next) =>
                {
                    var req = environment.Request;
                    string endpoint = environment.Request.Path.ToString();
                    Uri uri = req.Uri;
                    var seg = uri.Segments;

                    string verb = environment.Request.Method;
                    int userid = 0;
                    Int32.TryParse(environment.Request.Headers.Get("id"),out userid);
                    string token = environment.Request.Headers.Get("token");

                    ValidateAuth validator = new ValidateAuth();
                    int resourceid = 0;
                    //tiene resourseid
                    if (Int32.TryParse(seg[seg.Length-1], out resourceid))
                    {
                        endpoint = "";
                        for (int i = 0; i < seg.Length-1; i++)
                        {
                            endpoint += seg[i];
                        }
                    }

                    if (!validator.shallYouPass(userid, token, endpoint, verb) && !debugmode)
                    {
                        environment.Response.StatusCode = 401;
                        environment.Response.Body = new MemoryStream();

                        var newBody = new MemoryStream();
                        newBody.Seek(0, SeekOrigin.Begin);
                        var newContent = new StreamReader(newBody).ReadToEnd();

                        newContent += "You shall no pass.";

                        environment.Response.Body = newBody;
                        environment.Response.Write(newContent);
                    }
                    else
                    await next();
                }
            );
            ConfigureAuth(app);
        }
    }
}
