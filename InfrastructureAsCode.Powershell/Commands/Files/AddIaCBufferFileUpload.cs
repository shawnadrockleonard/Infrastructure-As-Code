using InfrastructureAsCode.Core;
using InfrastructureAsCode.Core.Services;
using InfrastructureAsCode.Powershell.CmdLets;
using InfrastructureAsCode.Powershell.PipeBinds;
using Microsoft.SharePoint.Client;
using System;
using System.IO;
using System.Management.Automation;

namespace InfrastructureAsCode.Powershell.Commands.Files
{
    /// <summary>
    /// The following command will upload a file to the specified folder.  
    ///     If not folder is specified it will upload the file to the root folder.
    /// </summary>
    /// <remarks>
    /// You can find documentation at https://github.com/pinch-perfect/Infrastructure-As-Code/blob/master/HowToExtend/Gist/sample-add-iac-buffer-fileupload.md
    /// </remarks>
    [Cmdlet(VerbsCommon.Add, "IaCBufferFileUpload")]
    [CmdletHelp("Uploads a large file via chunking to a library specified", Category = "Files")]
    public class AddIaCBufferFileUpload : IaCCmdlet
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
                throw new System.IO.FileNotFoundException(String.Format("{0} {1}", this.FileName, InfrastructureAsCode.Core.Properties.Resources.FileDoesNotExist));
            }
        }

        /// <summary>
        /// Execute the command uploading the file to the root of the document library
        /// </summary>
        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();

            var ilogger = new DefaultUsageLogger(
                (string arg1, object[] arg2) =>
                {
                    LogVerbose(arg1, arg2);
                },
                (string arg1, object[] arg2) =>
                {
                    LogWarning(arg1, arg2);
                },
                (Exception ex, string arg1, object[] arg2) =>
                {
                    LogError(ex, arg1, arg2);
                }
            );

            try
            {
                if (string.IsNullOrEmpty(FolderName))
                {
                    FolderName = "";
                }

                ilogger.LogInformation("Entering Library Upload Cmdlet");
                var ctx = this.ClientContext;

                var accessToken = ctx.GetAccessToken();

                var w = ctx.Web;
                var l = ListTitle.GetList(ctx.Web);
                ctx.Load(w, wctx => wctx.Url, wctx => wctx.ServerRelativeUrl);
                ctx.Load(l, listctx => listctx.RootFolder, listctx => listctx.RootFolder.ServerRelativeUrl);
                ClientContext.ExecuteQueryRetry();

                var webUri = new Uri(w.Url);
                var serverRelativeUrl = new Uri(webUri, Path.Combine(l.RootFolder.ServerRelativeUrl, FolderName));
                ilogger.LogInformation("Uploading file to {0}", serverRelativeUrl.AbsoluteUri);

                var fileService = new FileUploadService(ctx, ilogger);

                if (fileService.UploadFileWithBuffer(l, serverRelativeUrl, FileName, 8))
                {
                    ilogger.LogInformation("Successfully uploaded {0}", FileName);
                }

            }
            catch (Exception ex)
            {
                ilogger.LogError(ex, "Failed in SetFileUpload File:{0}", this.FileName);
            }

        }




    }
}
