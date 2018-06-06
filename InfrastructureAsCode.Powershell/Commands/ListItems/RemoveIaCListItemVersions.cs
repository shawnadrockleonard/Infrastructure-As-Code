using InfrastructureAsCode.Powershell.CmdLets;
using Microsoft.SharePoint.Client;
using System;
using System.Linq;
using System.Management.Automation;

namespace InfrastructureAsCode.Powershell.Commands.ListItems
{
    /// <summary>
    /// Queries list for items with versions and removes previous versions
    /// </summary>
    [Cmdlet(VerbsCommon.Remove, "IaCListItemVersions", SupportsShouldProcess = true)]
    public class RemoveIaCListItemVersions : IaCCmdlet
    {
        [Parameter(Mandatory = true)]
        public string LibraryName { get; set; }

        /// <summary>
        /// Execute the cmdlet
        /// </summary>
        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();
            try
            {
                var ctx = this.ClientContext;

                var ctxWeb = ctx.Web;
                var ctxList = ctxWeb.Lists.GetByTitle(LibraryName);
                ctx.Load(ctxWeb, w => w.ServerRelativeUrl);
                ctx.Load(ctxList, l => l.EnableVersioning, l => l.ItemCount, l => l.Id, l => l.BaseTemplate, l => l.BaseType, l => l.OnQuickLaunch, l => l.DefaultViewUrl, l => l.Title, l => l.Hidden);
                this.ClientContext.ExecuteQueryRetry();

                var itemCount = ctxList.ItemCount;
                LogVerbose(string.Format("The library {0} has {1} items", LibraryName, itemCount));

                if (itemCount > 0)
                {
                    var webRelativeUrl = ctxWeb.ServerRelativeUrl;
                    CamlQuery query = CamlQuery.CreateAllItemsQuery(150, new string[] { "Id", "CheckOutUser", "FileRef" });
                    ListItemCollectionPosition itemPosition = null;

                    while (true)
                    {
                        query.ListItemCollectionPosition = itemPosition;
                        var queryListItems = ctxList.GetItems(query);
                        this.ClientContext.Load(queryListItems);
                        this.ClientContext.ExecuteQueryRetry(1, 100);
                        itemPosition = queryListItems.ListItemCollectionPosition;

                        foreach (var listItem in queryListItems)
                        {
                            if (listItem.FileSystemObjectType == FileSystemObjectType.File)
                            {
                                var fileRef = listItem["FileRef"].ToString();
                                LogVerbose("Verify if there are versions for File {0}", fileRef);
                                var listItemRelativeUrl = string.Format("{0}/{1}/{2}_.000", webRelativeUrl, ctxList.Title, listItem.Id);
                                var ctxRelativeUrl = ctxWeb.GetFileByServerRelativeUrl(fileRef);
                                this.ClientContext.Load(ctxRelativeUrl, file => file.ListItemAllFields, file => file.Versions);
                                this.ClientContext.ExecuteQueryRetry();

                                if (ctxRelativeUrl.Versions != null && ctxRelativeUrl.Versions.Any())
                                {
                                    foreach (var version in ctxRelativeUrl.Versions)
                                    {
                                        LogVerbose("Version: {0} Is Current:{2} with size:{1}", version.VersionLabel, version.Size, version.IsCurrentVersion);
                                    }

                                    if (this.ShouldProcess(string.Format("Deleting version history for {0}", fileRef)))
                                    {
                                        // Delete Versions
                                        ctxRelativeUrl.Versions.DeleteAll();
                                        this.ClientContext.ExecuteQuery();
                                    }
                                }
                            }
                        }
                        if (itemPosition == null)
                        {
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogError(ex, "Failed in GetListItemCount for Library {0}", LibraryName);
            }
        }
    }
}
