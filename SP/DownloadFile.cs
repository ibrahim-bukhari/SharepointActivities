using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Activities;
using System.ComponentModel;
using System.Security;
using Microsoft.SharePoint.Client;
using System.IO;

namespace SP
{
    public class DownloadFile : CodeActivity 
    {
        [Category("Input")]
        [RequiredArgument]
        public InArgument<String> URL { get; set; }

        [Category("Input")]
        [RequiredArgument]
        public InArgument<String> Username { get; set; }

        [Category("Input")]
        [RequiredArgument]
        public InArgument<SecureString> Password { get; set; }


        [Category("Input")]
        [RequiredArgument]
        public InArgument<String> OutputPath { get; set; }

        [Category("Output")]
        public OutArgument<Exception> Exception { get; set; }

        protected override void Execute(CodeActivityContext context)
        {

            ClientContext client = null;
            try
            {
                string url = context.GetValue(this.URL);
                string username = context.GetValue(this.Username);
                SecureString password = context.GetValue(this.Password);
                string output = context.GetValue(this.OutputPath);

                //var credentials = new SharePointOnlineCredentials(username, password);


                Uri filename = new Uri(@url);
                string server = filename.AbsoluteUri.Replace(filename.AbsolutePath, "");
                string serverrelative = filename.AbsolutePath;


                
                client = Helper.Auth(username, password, url);
                if (client != null)
                {
                    FileInformation f = Microsoft.SharePoint.Client.File.OpenBinaryDirect(client, serverrelative);
                    client.ExecuteQuery();

                    using (var fileStream = new FileStream(@output, FileMode.Create))
                        f.Stream.CopyTo(fileStream);
                }
                else
                {
                    throw new System.Exception("Couldn't establish connection to SharePoint Site");
                }

            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
                Exception.Set(context, ex);
                throw ex;
            }
            finally
            {
                if(client != null)
                {
                    client.Dispose();
                }
            }
            
        }
    }
}
