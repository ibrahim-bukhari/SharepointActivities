using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace SP
{
    class Helper
    {
        public static ClientContext Auth(String uname, SecureString pwd, string siteURL)
        {
            ClientContext context = new ClientContext(siteURL);
            Web web = context.Web;
            context.Credentials = new SharePointOnlineCredentials(uname, pwd);
            try
            {
                context.Load(web);
                context.ExecuteQuery();
                System.Diagnostics.Debug.WriteLine("Hello! from " + web.Title + " site");
                return context;
            }
            catch (Exception e)
            {
                Console.WriteLine("Something went wrong in SP Auth Module" + e);
                return null;
            }
        }
    }
}
