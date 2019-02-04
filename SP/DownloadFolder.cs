using Microsoft.SharePoint.Client;
using System;
using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace SP
{
    public class DownloadFolder : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        public InArgument<String> URL { get; set; }

        [Category("Input")]
        [RequiredArgument]
        public InArgument<String> Library { get; set; }

        [Category("Input")]
        public InArgument<String> Folder { get; set; }

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

            try
            {
                string url = context.GetValue(this.URL);
                string username = context.GetValue(this.Username);
                SecureString password = context.GetValue(this.Password);
                string library = context.GetValue(this.Library);
                string folder = context.GetValue(this.Folder);
                string output = context.GetValue(this.OutputPath);

                DownloadFiles(username, password, url, library, folder, output);
               
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Exception.Set(context, ex);
                throw ex;
            }

        }


        private static void DownloadFiles(String username, SecureString password, string url, string library, string folder, string output)
        {
            if(string.IsNullOrEmpty(username))
            {
                throw new System.Exception("Username not specified");
            }

            if (string.IsNullOrEmpty(url))
            {
                throw new System.Exception("Sharepoint URL not specified");
            }

            if (string.IsNullOrEmpty(library))
            {
                throw new System.Exception("Sharepoint Library not specified");
            }

            if(string.IsNullOrEmpty(output))
            {
                throw new System.Exception("Output location not specified");
            }

            if (string.IsNullOrEmpty(folder))
            {
                Console.WriteLine("Sharepoint folder name not specified. Use root folder");
            }

            ClientContext client = null;
            try
            {
                client = Helper.Auth(username, password, url);
                if (client != null)
                {
                    var web = client.Web;
                    client.Load(web);
                    client.ExecuteQuery();

                    List list = web.Lists.GetByTitle(library);
                    client.Load(list);
                    client.ExecuteQuery();
                    client.Load(list.RootFolder);
                    client.ExecuteQuery();
                    client.Load(list.RootFolder.Folders);
                    client.ExecuteQuery();
                    processFolderClientobj(list.RootFolder.ServerRelativeUrl + folder, output, client);
                    //foreach (Folder f in list.RootFolder.Folders)
                    //{
                    //    processFolderClientobj(f.ServerRelativeUrl);
                    //}
                }
                else
                {
                    throw new System.Exception("Couldn't establish connection to SharePoint Site");
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw (ex);
            }
            finally
            {
                if(client != null)
                    client.Dispose();
            }
        }

        private static void processFolderClientobj(string folderURL, string Destination, ClientContext site)
        {
            var web = site.Web;
            site.Load(web);
            site.ExecuteQuery();
            Console.WriteLine("Sharepoint folder path: " + folderURL);
            Folder folder = web.GetFolderByServerRelativeUrl(folderURL);
            site.Load(folder);
            site.ExecuteQuery();
            site.Load(folder.Files);
            site.ExecuteQuery();
            int filenum = 1;

            foreach (Microsoft.SharePoint.Client.File file in folder.Files)
            {
                int numberOfFiles = folder.Files.Count;
                string destinationfolder = Destination + "/" + folder.ServerRelativeUrl;
                Stream fs = Microsoft.SharePoint.Client.File.OpenBinaryDirect(site, file.ServerRelativeUrl).Stream;
                byte[] binary = ReadFully(fs);
                if (!Directory.Exists(destinationfolder))
                {
                    Directory.CreateDirectory(destinationfolder);
                }

                string filename = file.Name;
                Console.WriteLine("Downloading file " + filenum + " of " + numberOfFiles + " - " + filename);
                FileStream stream = new FileStream(destinationfolder + "/" + filename, FileMode.Create);
                BinaryWriter writer = new BinaryWriter(stream);
                writer.Write(binary);
                writer.Close();
                filenum++;
            }
        }

        private static byte[] ReadFully(Stream input)
        {
            byte[] buffer = new byte[16 * 1024];
            using (MemoryStream ms = new MemoryStream())
            {
                int read;
                while ((read = input.Read(buffer, 0, buffer.Length)) > 0)
                {
                    ms.Write(buffer, 0, read);
                }
                return ms.ToArray();
            }
        }

    }
}
