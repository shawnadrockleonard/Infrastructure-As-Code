using InfrastructureAsCode.Core.Models;
using InfrastructureAsCode.Powershell.CmdLets;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Utilities;
using System;
using System.Collections.Generic;
using System.Management.Automation;
using System.Linq;

namespace InfrastructureAsCode.Powershell.Commands.ETL
{
    /// <summary>
    /// The function cmdlet will upgrade the site specified in the connection to the latest configuration changes
    /// </summary>
    [Cmdlet(VerbsCommon.Set, "IaCMigrateData", SupportsShouldProcess = true)]
    [CmdletHelp("For a list that might require external process, a DataMigrated column stores if the record has been processed", Category = "ETL")]
    public class SetIaCMigrateData : IaCCmdlet
    {
        /// <summary>
        /// Represents the title of the list/library
        /// </summary>
        [Parameter(Mandatory = false)]
        public string ListTitle { get; set; }

        /// <summary>
        /// The single SiteAsset file to upload based on relative path
        /// </summary>
        [Parameter(Mandatory = false)]
        public string ActionFile { get; set; }

        /// <summary>
        /// Validate parameters
        /// </summary>
        protected override void OnBeginInitialize()
        {
            if (!string.IsNullOrEmpty(this.ActionFile) && !System.IO.File.Exists(this.ActionFile))
            {
                throw new Exception(string.Format("The file does not exists {0}", this.ActionFile));
            }
        }

        private List<SPFieldDefinitionModel> siteColumns { get; set; }

        /// <summary>
        /// Process the request
        /// </summary>
        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();

            siteColumns = new List<SPFieldDefinitionModel>();

            try
            {
                //Load list
                var siteUrl = this.ClientContext.Url;
                var listInSite = this.ClientContext.Web.Lists.GetByTitle(ListTitle);
                this.ClientContext.Load(listInSite);
                this.ClientContext.ExecuteQuery();

                //#TODO: Provision datamigrated column for holder
                var camlWhereClause = CAML.Neq(CAML.FieldValue("DataMigrated", "Integer", "1"));

                var camlViewFields = CAML.ViewFields((new string[] { "Modified", "", "", "Id" }).Select(s => CAML.FieldRef(s)).ToArray());

                // get site and query the list for approved requests
                ListItemCollectionPosition ListItemCollectionPosition = null;
                var camlQuery = new CamlQuery()
                {
                    ViewXml = CAML.ViewQuery(ViewScope.RecursiveAll, CAML.Where(camlWhereClause), string.Empty, camlViewFields, 50)
                };


                var output = new List<object>();

                while (true)
                {
                    camlQuery.ListItemCollectionPosition = ListItemCollectionPosition;
                    var spListItems = listInSite.GetItems(camlQuery);
                    this.ClientContext.Load(spListItems, lti => lti.ListItemCollectionPosition,
                        lti => lti.IncludeWithDefaultProperties(lnc => lnc.Id, lnc => lnc.ContentType));
                    this.ClientContext.ExecuteQuery();
                    ListItemCollectionPosition = spListItems.ListItemCollectionPosition;

                    foreach (var requestItem in spListItems)
                    {
                        var requestId = requestItem.Id;

                        ListItem _item = listInSite.GetItemById(requestId);
                        ClientContext.Load(_item);
                        ClientContext.ExecuteQueryRetry();

                        try
                        {
                            output.Add(ProcessListItem(_item));
                        }
                        catch (Exception e)
                        {
                            LogError(e, "Failed to update list item {0}", e.Message);
                        }
                    }

                    if (ListItemCollectionPosition == null)
                    {
                        break;
                    }

                }

                LogVerbose("Writing objects to memory stream.");
                output.ForEach(s => WriteObject(s));
            }
            catch (Exception ex)
            {
                LogError(ex, "Migrate failed for list items MSG:{0}", ex.Message);
            }
        }

        /// <summary>
        /// Provides an override method for modifying the list item
        /// </summary>
        /// <param name="_item">The ListItem on which the operation will migrate data</param>
        protected virtual object ProcessListItem(ListItem _item)
        {
            // Retain original field settings
            // #TODO check based on list content types
            var modifiedDate = _item["Modified"];
            var modifiedBy = _item["Editor"];

            if (this.ShouldProcess(string.Format("update list basic properties {0}", _item.Id)))
            {
                _item["DataMigrated"] = 1;
                _item.SystemUpdate();
                ClientContext.ExecuteQueryRetry();

                try
                {
                    _item["Modified"] = modifiedDate;
                    _item["Editor"] = modifiedBy;
                    _item.SystemUpdate();
                    ClientContext.ExecuteQueryRetry();
                }
                catch (Exception e)
                {
                    LogError(e, "Failed to update list basic properties {0}", e.Message);
                }
            }
            return null;
        }

    }
}
