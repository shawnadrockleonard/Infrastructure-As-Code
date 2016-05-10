using InfrastructureAsCode.Powershell;
using InfrastructureAsCode.Powershell.CmdLets;
using InfrastructureAsCode.Powershell.Extensions;
using InfrastructureAsCode.Powershell.PipeBinds;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Powershell.Commands.Files
{
    [Cmdlet(VerbsCommon.Add, "IaCFileUpload")]
    [CmdletHelp("Uploads a document or file to a library specified", Category = "Files")]
    public class AddIaCFileUpload : IaCCmdlet
    {
        /// <summary>
        /// The display name for the library
        /// </summary>
        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
        public string ListTitle { get; set; }

        /// <summary>
        /// The full path to the file
        /// </summary>
        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 1)]
        public string FileName { get; set; }


        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            if (!System.IO.File.Exists(this.FileName))
            {
                throw new InvalidOperationException(String.Format("{0} {1}", this.FileName, InfrastructureAsCode.Core.Properties.Resources.FileDoesNotExist));
            }
        }

        /// <summary>
        /// Execute the command uploading the file to the root of the document library
        /// </summary>
        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();
            try
            {
                LogVerbose("Entering Library Upload Cmdlet");
                var ctx = this.ClientContext;

                var w = ctx.Web;
                var l = ctx.Web.Lists.GetByTitle(ListTitle);
                ctx.Load(w);
                ctx.Load(l, listEntity => listEntity.RootFolder.ServerRelativeUrl);
                ClientContext.ExecuteQueryRetry();

                var serverRelativeUrl = string.Format("{0}/{1}", this.ClientContext.Url, l.RootFolder.ServerRelativeUrl);
                LogVerbose(string.Format("Context has been established for {0}", serverRelativeUrl));

                var fileName = Path.GetFileName(this.FileName);
                using (var stream = new System.IO.FileStream(this.FileName, FileMode.Open))
                {

                    var creationInfo = new Microsoft.SharePoint.Client.FileCreationInformation();
                    creationInfo.Overwrite = true;
                    creationInfo.ContentStream = stream;
                    creationInfo.Url = fileName;

                    var uploadStatus = l.RootFolder.Files.Add(creationInfo);
                    ctx.Load(uploadStatus);

                    ctx.ExecuteQuery();
                }
            }
            catch (Exception ex)
            {
                LogError(ex, "Failed in SetFileUpload File:{0}", this.FileName);
            }

        }
    }
}
