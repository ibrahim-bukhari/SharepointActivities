using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Activities;
using System.ComponentModel;
using System.Security;
using Microsoft.SharePoint.Client;
using System.IO;

namespace SP
{

    public class UploadFolder : CodeActivity
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
        public InArgument<String> InputPath { get; set; }

        [Category("Output")]
        public OutArgument<Exception> Exception { get; set; }

        // If your activity returns a value, derive from CodeActivity<TResult>
        // and return the value from the Execute method.
        protected override void Execute(CodeActivityContext context)
        {
            try
            {
                // Obtain the runtime value of input arguments
                string url = context.GetValue(this.URL);
                string username = context.GetValue(this.Username);
                SecureString password = context.GetValue(this.Password);
                string library = context.GetValue(this.Library);
                string folder = context.GetValue(this.Folder);
                string input = context.GetValue(this.InputPath);

                UploadFiles(username, password, url, library, folder, input);
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
                Exception.Set(context, ex);
                throw ex;
            }
        }

        private static void UploadFiles(string username, SecureString password, string url, string library, string folder, string input)
        {
            if (string.IsNullOrEmpty(username))
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

            if (string.IsNullOrEmpty(input))
            {
                throw new System.Exception("Input location not specified");
            }

            if (string.IsNullOrEmpty(folder))
            {
                Console.WriteLine("Sharepoint folder name not specified. Use root folder");
            }

            ClientContext client = null;
            try
            {
                client = Helper.Auth(username, password, url);
                //Assume that the web site has a library named "FormLibrary".
                if (client != null)
                {
                    var formLib = client.Web.Lists.GetByTitle(library);
                    client.Load(formLib.RootFolder);
                    client.ExecuteQuery();

                    string[] files = null;
                    if (Directory.Exists(input))
                    {
                        files = Directory.GetFiles(input, @"*.*", SearchOption.TopDirectoryOnly);
                    }
                    int numberOfFiles = files.Length;
                    if (numberOfFiles > 0)
                    {
                        for (int i = 0; i < numberOfFiles; i++)
                        {
                            string file = files[i];
                            if (System.IO.File.Exists(file) && !file.Contains('~'))
                            {
                                var fileUrl = "";
                                //Craete FormTemplate and save in the library.
                                using (var fs = new FileStream(file, FileMode.Open))
                                {
                                    string filename = Path.GetFileName(file);
                                    int filenum = i + 1;
                                    Console.WriteLine("Uploading file " + filenum + " of " + numberOfFiles + " - " + filename);
                                    var fi = new FileInfo(filename);
                                    fileUrl = String.Format("{0}{1}{2}", formLib.RootFolder.ServerRelativeUrl, folder, fi.Name);
                                    Console.WriteLine(fileUrl);
                                    Microsoft.SharePoint.Client.File.SaveBinaryDirect(client, fileUrl, fs, true);
                                    client.ExecuteQuery();
                                }
                            }
                        }
                        Console.WriteLine("All files uploaded");
                    }
                    else
                    {
                        Console.WriteLine("No files to upload");
                    }

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
                if (client != null)
                    client.Dispose();
            }
        }

    }
}
