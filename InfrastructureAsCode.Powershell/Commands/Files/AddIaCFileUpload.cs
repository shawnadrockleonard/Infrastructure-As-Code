using InfrastructureAsCode.Powershell;
using InfrastructureAsCode.Powershell.Commands.Base;
using InfrastructureAsCode.Core.Extensions;
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
        #region Parameters 

        /// <summary>
        /// The display name for the library
        /// </summary>
        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
        public ListPipeBind ListTitle { get; set; }

        /// <summary>
        /// The full path to the file
        /// </summary>
        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 1)]
        public string FileName { get; set; }

        /// <summary>
        /// The foldername in which the file will be uploaded
        /// </summary>
        [Parameter(Mandatory = false, ValueFromPipeline = true, Position = 2)]
        public string FolderName { get; set; }

        /// <summary>
        /// Should we overwrite the file if it exists
        /// </summary>
        [Parameter(Mandatory = false, ValueFromPipeline = true, Position = 3)]
        public SwitchParameter Clobber { get; set; }

        #endregion


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
                var l = ListTitle.GetList(ctx.Web);
                ctx.Load(w, wctx => wctx.Url, wctx => wctx.ServerRelativeUrl);
                ctx.Load(l, listctx => listctx.RootFolder, listctx => listctx.RootFolder.ServerRelativeUrl);
                ClientContext.ExecuteQueryRetry();

                var webUri = new Uri(w.Url);
                var serverRelativeUrl = new Uri(webUri, Path.Combine(l.RootFolder.ServerRelativeUrl, FolderName));
                LogVerbose(string.Format("Uploading file to {0}", serverRelativeUrl.AbsoluteUri));

                if (string.IsNullOrEmpty(FolderName))
                {
                    var uploadedFileUrl = l.UploadFile(l.RootFolder, FileName, Clobber);
                    LogVerbose(string.Format("Uploaded [Clobber:{1}] into RootFolder file URL {0}", uploadedFileUrl, Clobber));
                }
                else
                {
                    var uploadedFileUrl = l.UploadFile(FolderName, FileName, Clobber);
                    LogVerbose(string.Format("Uploaded [Clobber:{2}] into Folder {0} file URL {1}", FolderName, uploadedFileUrl, Clobber));
                }

            }
            catch (Exception ex)
            {
                LogError(ex, "Failed in SetFileUpload File:{0}", this.FileName);
            }

        }
    }
}
