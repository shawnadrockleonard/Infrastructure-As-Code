using InfrastructureAsCode.Core.Models;
using InfrastructureAsCode.Powershell.CmdLets;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Utilities;
using System;
using System.Collections.Generic;
using System.Management.Automation;

namespace InfrastructureAsCode.Powershell.Commands
{
    /// <summary>
    /// The function cmdlet will upgrade the EzForms site specified in the connection to the latest configuration changes
    /// </summary>
    [Cmdlet(VerbsCommon.Set, "IaCMigrateData")]
    [CmdletHelp("Identify users via json file and send email", Category = "ETL")]
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
        protected override void BeginProcessing()
        {
            base.BeginProcessing();
            if (!string.IsNullOrEmpty(this.ActionFile) && !System.IO.File.Exists(this.ActionFile))
            {
                throw new Exception(string.Format("The file does not exists {0}", this.ActionFile));
            }
        }

        internal List<SPFieldDefinitionModel> siteColumns { get; set; }

        /// <summary>
        /// Process the request
        /// </summary>
        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();

            siteColumns = new List<SPFieldDefinitionModel>();


            if (this.ClientContext == null)
            {
                LogWarning("Invalid client context, configure the service to run again");
                return;
            }

            try
            {
                //Load list
                var siteUrl = this.ClientContext.Url;
                var listInSite = this.ClientContext.Web.Lists.GetByTitle(ListTitle);
                this.ClientContext.Load(listInSite);
                this.ClientContext.ExecuteQuery();

                //#TODO: Provision datamigrated column for holder
                var camlWhereClause = CAML.Neq(CAML.FieldValue("DataMigrated", "Integer", "1"));

                // get ezforms site and query the list for approved requests
                ListItemCollectionPosition ListItemCollectionPosition = null;
                var camlQuery = CamlQuery.CreateAllItemsQuery();
                camlQuery.ViewXml = string.Format("<View Scope='RecursiveAll'><Query><Where>{0}</Where>", camlWhereClause);
                camlQuery.ViewXml += "<RowLimit>50</RowLimit>";
                camlQuery.ViewXml += "</Query></View>";
                camlQuery.ListItemCollectionPosition = ListItemCollectionPosition;

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
                        //#TODO: add list item column includes for specific columns of operation
                        ClientContext.Load(_item);
                        ClientContext.ExecuteQuery();

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

            if (!DoNothing)
            {
                _item["DataMigrated"] = 1;
                _item.Update();
                ClientContext.ExecuteQueryRetry();

                try
                {
                    _item["Modified"] = modifiedDate;
                    _item["Editor"] = modifiedBy;
                    _item.Update();
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
