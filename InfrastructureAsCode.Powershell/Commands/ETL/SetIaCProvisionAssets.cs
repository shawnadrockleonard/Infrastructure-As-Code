using InfrastructureAsCode.Powershell;
using InfrastructureAsCode.Powershell.CmdLets;
using InfrastructureAsCode.Core.Models;
using InfrastructureAsCode.Powershell.Extensions;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using OfficeDevPnP.Core.Entities;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace InfrastructureAsCode.Powershell.Commands.ETL
{
    /// <summary>
    /// The function cmdlet will push local files into the Site Assets
    /// </summary>
    [Cmdlet(VerbsCommon.Set, "IaCProvisionAssets")]
    [CmdletHelp("Push files into the site assets library.", Category = "ETL")]
    public class SetIaCProvisionAssets : IaCCmdlet
    {
        /// <summary>
        /// Represents the directory path for any files for serialization
        /// </summary>
        [Parameter(Mandatory = false)]
        public string SiteContent { get; set; }

        /// <summary>
        /// The single SiteAsset file to upload based on relative path
        /// </summary>
        [Parameter(Mandatory = false)]
        public string SiteActionFile { get; set; }

        /// <summary>
        /// Validate parameters
        /// </summary>
        protected override void BeginProcessing()
        {
            base.BeginProcessing();
            if (!System.IO.Directory.Exists(this.SiteContent))
            {
                throw new Exception(string.Format("The directory does not exists {0}", this.SiteContent));
            }
        }

        /// <summary>
        /// Process the request
        /// </summary>
        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();

            if (this.ClientContext == null)
            {
                LogWarning("Invalid client context, configure the service to run again");
                return;
            }

            // obtain CSOM object for host web
            Web hostWeb = this.ClientContext.Web;

            // check to see if Picture library named Photos already exists
            ListCollection allLists = hostWeb.Lists;
            IEnumerable<List> foundLists = this.ClientContext.LoadQuery(allLists.Where(list => list.Title == "Site Assets"));
            this.ClientContext.ExecuteQuery();
            List siteAssetsLibrary = foundLists.FirstOrDefault();
            if (siteAssetsLibrary == null)
            {
                LogWarning("Failed to find site assets");
                return;
            }

            var appFileFolder = string.Format("{0}\\SiteAssets", this.SiteContent);
            if (!System.IO.Directory.Exists(appFileFolder))
            {
                LogWarning("Site Assets Folder {0} does not exist.", appFileFolder);
                return;
            }

            var searchPattern = "*";
            if (!string.IsNullOrEmpty(this.SiteActionFile))
            {
                searchPattern = this.SiteActionFile;
                if (this.SiteActionFile.IndexOf(@"\") > -1)
                {
                    searchPattern = this.SiteActionFile.Substring(0, this.SiteActionFile.IndexOf(@"\"));
                }
            }

            var appDirectories = System.IO.Directory.GetDirectories(appFileFolder, searchPattern, System.IO.SearchOption.TopDirectoryOnly);
            foreach (var appDirectory in appDirectories)
            {
                var appDirectoryInfo = new System.IO.DirectoryInfo(appDirectory);

                string folder = appDirectoryInfo.Name, folderPath = appDirectoryInfo.FullName;

                var searchPatternFiles = "*";
                if (!string.IsNullOrEmpty(this.SiteActionFile))
                {
                    if (this.SiteActionFile.IndexOf(@"\") > -1)
                    {
                        searchPatternFiles = string.Format("*{0}*", this.SiteActionFile.Substring(this.SiteActionFile.IndexOf(@"\") + 1));
                    }
                }

                var appRootFolder = siteAssetsLibrary.RootFolder.EnsureFolder(folder);
                var siteAssetsFiles = System.IO.Directory.GetFiles(folderPath, searchPatternFiles);
                LogVerbose("Now searching folder {0} and uploading files {1}", folderPath, siteAssetsFiles.Count());

                // enmuerate through each file in folder
                foreach (string filePath in siteAssetsFiles)
                {
                    LogVerbose("---------------- Now uploading file {0}", filePath);

                    if (!DoNothing)
                    {
                        // upload each file to library in host web
                        byte[] fileContent = System.IO.File.ReadAllBytes(filePath);
                        FileCreationInformation fileInfo = new FileCreationInformation();
                        fileInfo.Content = fileContent;
                        fileInfo.Overwrite = true;
                        fileInfo.Url = System.IO.Path.GetFileName(filePath);
                        File newFile = appRootFolder.Files.Add(fileInfo);

                        // commit changes to library
                        this.ClientContext.Load(newFile, nf => nf.ServerRelativeUrl, nf => nf.Length);
                        this.ClientContext.ExecuteQuery();
                    }
                }
            }
        }


    }
}
