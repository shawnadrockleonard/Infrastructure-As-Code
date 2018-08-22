using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Powershell.Commands.ListItems
{
    using Microsoft.SharePoint.Client;
    using InfrastructureAsCode.Powershell.PipeBinds;
    using InfrastructureAsCode.Powershell.Commands.Base;
    using InfrastructureAsCode.Core.Models;
    using InfrastructureAsCode.Powershell;
    using OfficeDevPnP.Core.Utilities;

    /// <summary>
    /// This command will find versions for list items and remove those versions
    /// </summary>
    [Cmdlet(VerbsCommon.Remove, "IaCListItemVersions", SupportsShouldProcess = true)]
    [CmdletHelp("Queries list for items with versions and removes previous versions", Category = "ListItems")]
    public class RemoveIaCListItemVersions : IaCCmdlet
    {
        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
        public ListPipeBind ListTitle { get; set; }

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
                ctx.Load(ctxWeb, w => w.ServerRelativeUrl);
                this.ClientContext.ExecuteQueryRetry();

                var ctxList = ListTitle.GetList(ctxWeb,
                    l => l.EnableVersioning,
                    l => l.ItemCount,
                    l => l.Id,
                    l => l.BaseTemplate,
                    l => l.BaseType,
                    l => l.OnQuickLaunch,
                    l => l.DefaultViewUrl,
                    l => l.Title, l => l.Hidden);


                var itemCount = ctxList.ItemCount;
                LogVerbose(string.Format("The library {0} has {1} items", ListTitle, itemCount));

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

                                    if (this.ShouldProcess(string.Format("---------------- Now deleting {0} versions", ctxRelativeUrl.Versions.Count())))
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
                LogError(ex, "Failed in RemoveIaCListItemVersions for Library {0}", ListTitle);
            }
        }
    }
}
