using InfrastructureAsCode.Powershell.CmdLets;
using InfrastructureAsCode.Powershell.PipeBinds;
using InfrastructureAsCode.Core.Extensions;
using OfficeDevPnP.Core.Extensions;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Powershell.Commands.Lists
{
    /// <summary>
    /// Adds folders to a sharepoint list
    /// </summary>
    /// <remarks>
    /// Add-IaCListFolders -Identity /Lists/Announcements
    /// Add-IaCListFolders -Identity 99a00f6e-fb81-4dc7-8eac-e09c6f9132fe"
    /// </remarks>
    [Cmdlet(VerbsCommon.Add, "IaCListFolders")]
    public class AddIaCListFolders : IaCCmdlet
    {
        #region Parameters

        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0, HelpMessage = "The ID or Url of the list.")]
        public ListPipeBind Identity { get; set; }

        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 1, HelpMessage = "String separated folder list")]
        public string FolderPath { get; set; }

        #endregion

        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();
            try
            {
                LogVerbose("Entering Folder creation Cmdlet");
                var ctx = this.ClientContext;

                var w = ctx.Web;
                var l = Identity.GetList(ctx.Web);
                ctx.Load(w, wctx => wctx.Url, wctx => wctx.ServerRelativeUrl);
                ctx.Load(l, lctx => lctx.ItemCount, lctx => lctx.EnableFolderCreation, lctx => lctx.RootFolder, lctx => lctx.RootFolder.ServerRelativeUrl, lctx => lctx.RootFolder.Folders);
                ctx.ExecuteQueryRetry();

                var webUri = new Uri(w.Url);

                var folderNames = FolderPath.Split(new string[] { "/", "\\" }, StringSplitOptions.RemoveEmptyEntries);
                var currentFolder = l.RootFolder;

                // Will query the list to determine the last item id in the list
                var lastItemId = l.QueryLastItemId();
                LogVerbose("List with item count {0} has a last ID of {1}", l.ItemCount, lastItemId);
                LogVerbose("List has folder creation = {0}", l.EnableFolderCreation);

                foreach (var foldername in folderNames)
                {
                    currentFolder = l.GetOrCreateFolder(currentFolder, foldername, 0, lastItemId);

                    if(!currentFolder.IsPropertyAvailable(fctx => fctx.ServerRelativeUrl))
                    {
                        currentFolder.Context.Load(currentFolder, fctx => fctx.ServerRelativeUrl);
                        currentFolder.Context.ExecuteQueryRetry();
                    }

                    var serverRelativeUrl = new Uri(webUri, currentFolder.ServerRelativeUrl);
                    LogVerbose("Folder created or exists at {0}", serverRelativeUrl.AbsoluteUri);
                }
            }
            catch (Exception ex)
            {
                LogError(ex, "Failed create folder structure:{0}", this.FolderPath);
            }

        }
    }
}
