using IaC.Core.Models;
using IaC.Powershell;
using IaC.Powershell.CmdLets;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Utilities;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;

namespace IaC.Powershell.Commands.RecycleBin
{
    [Cmdlet(VerbsCommon.Reset, "IaCRecycleBinItems", SupportsShouldProcess = true)]
    [CmdletHelp("Query the recycle bin for the specific path and restore.", Category = "RecycleBin")]
    public class ResetIaCRecycleBinItems : IaCCmdlet
    {
        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
        public string SitePathUrl { get; set; }

        [Parameter(Mandatory = false, ValueFromPipeline = true, Position = 1)]
        public Nullable<int> RowLimit { get; set; }

        [Parameter(Mandatory = false, ValueFromPipeline = true, Position = 2)]
        public Nullable<RecycleBinItemType> RecycleBinType { get; set; }

        [Parameter(Mandatory = false, ValueFromPipeline = true, Position = 3)]
        public Nullable<RecycleBinOrderBy> RecycleBinOrder { get; set; }

        /// <summary>
        /// Execute the cmdlet
        /// </summary>
        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();
            if (!RecycleBinType.HasValue)
            {
                RecycleBinType = RecycleBinItemType.Folder;
            }
            if (!RecycleBinOrder.HasValue)
            {
                RecycleBinOrder = RecycleBinOrderBy.DeletedBy;
            }
            if (!RowLimit.HasValue)
            {
                RowLimit = 200;
            }

            var results = new Collection<SPRecycleBinElement>();
            var failedrestores = new Collection<SPRecycleBinElement>();
            var restorepath = new Collection<SPRecycleBinRestoreModel>();

            // this should be valid based on pre authentication checks
            var _currentUserInProcess = CurrentUserName;
            LogVerbose("Started-QueryRecycleBin: User:{0} at {1}", _currentUserInProcess, DateTime.Now);

            var collSite = this.ClientContext.Site;
            var collWeb = this.ClientContext.Web;
            this.ClientContext.Load(collSite);
            this.ClientContext.Load(collWeb);
            this.ClientContext.ExecuteQueryRetry();


            if (this.SitePathUrl.IndexOf("/", StringComparison.CurrentCultureIgnoreCase) == 0)
            {
                SitePathUrl = SitePathUrl.Substring(1);
            }

            var rbiQueryCount = 0;
            var rbiCollectionCount = 0;
            var pagingLimit = RowLimit.Value;
            var paging = true;
            string pagingInfo = null;
            try
            {
                while (paging)
                {
                    LogVerbose("Paging Recycle bin {0} with paging sequence [{1}]", SitePathUrl, pagingInfo);

                    RecycleBinItemCollection recycleBinCollection = collSite.GetRecycleBinItems(pagingInfo, pagingLimit, false, RecycleBinOrder.Value, RecycleBinItemState.FirstStageRecycleBin);
                    this.ClientContext.Load(recycleBinCollection);
                    this.ClientContext.ExecuteQueryRetry();
                    paging = false;

                    if (recycleBinCollection != null && recycleBinCollection.Count > 0)
                    {
                        var rbiQuery = recycleBinCollection.Where(w => w.DirName.StartsWith(this.SitePathUrl, StringComparison.CurrentCultureIgnoreCase)
                            && w.ItemType == RecycleBinType).OrderBy(ob => ob.DirName);
                        rbiQueryCount = rbiQuery.Count();
                        rbiCollectionCount = recycleBinCollection.Count();
                        LogVerbose("Query resulted in {0} items", rbiQueryCount);

                        foreach (var rbiItem in rbiQuery)
                        {
                            var newitem = new SPRecycleBinRestoreModel()
                            {
                                Title = rbiItem.Title,
                                Id = rbiItem.Id,
                                FileSize = rbiItem.Size,
                                LeafName = rbiItem.LeafName,
                                DirName = rbiItem.DirName,
                                FileType = rbiItem.ItemType,
                                FileState = rbiItem.ItemState,
                                Author = rbiItem.AuthorName,
                                AuthorEmail = rbiItem.AuthorEmail,
                                DeletedBy = rbiItem.DeletedByName,
                                DeletedByEmail = rbiItem.DeletedByEmail,
                                Deleted = rbiItem.DeletedDate,
                                PagingInfo = pagingInfo
                            };
                            
                            restorepath.Add(newitem);
                        }

                        if (rbiCollectionCount >= pagingLimit)
                        {
                            var pageItem = recycleBinCollection[recycleBinCollection.Count() - 1];
                            var leafNameEncoded = HttpUtility.UrlPathEncode(pageItem.Title, true).Replace(".", "%2E");
                            var searchEncoded = HttpUtility.UrlPathEncode(pageItem.DirName, true);
                            pagingInfo = string.Format("id={0}&title={1}&searchValue={2}", pageItem.Id, leafNameEncoded, searchEncoded);
                            paging = true;
                        }
                    }
                    else
                    {
                        LogVerbose("The Recycle Bin is empty.");
                    }
                }

                if(restorepath.Count() > 0)
                {
                    var orderedrestore = restorepath.OrderBy(ob => ob.DirName);
                    LogVerbose("Restoring: {0} items for {1}", restorepath.Count, SitePathUrl);
                    foreach (var newitem in orderedrestore)
                    {
                        if (!DoNothing)
                        {
                            try
                            {
                                var rbiItem = collSite.RecycleBin.GetById(newitem.Id);
                                this.ClientContext.Load(rbiItem);
                                this.ClientContext.ExecuteQueryRetry();

                                rbiItem.Restore();
                                this.ClientContext.ExecuteQueryRetry();
                                results.Add(newitem);
                                LogVerbose("Restored: {0}|{1}, Deleted By:{2}", newitem.DirName, newitem.LeafName, newitem.DeletedByEmail);
                            }
                            catch (Exception ex)
                            {
                                rbiQueryCount--;
                                failedrestores.Add(newitem);
                                LogWarning("Failed: {0}|{1} MSG:{2}", newitem.DirName, newitem.LeafName, ex.Message);
                            }
                        }
                    }
                }


            }
            catch (Exception ex)
            {
                LogError(ex,  "Failed to retrieve recycle bin collection");
            }

            WriteObject(results);
        }
    }
}
