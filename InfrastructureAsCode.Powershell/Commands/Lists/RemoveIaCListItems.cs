using InfrastructureAsCode.Powershell;
using InfrastructureAsCode.Powershell.CmdLets;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Powershell.Commands.Lists
{
    /// <summary>
    /// Demonstrates how to delete specific items based on the query specified. 
    ///     You can't modified an collection that is being enumerated so this presents a pattern to load up and process.
    /// </summary>
    [Cmdlet(VerbsCommon.Remove, "IaCListItems", SupportsShouldProcess = true)]
    [CmdletHelp("Query the specific list and delete items, if begin/end is specified filter the query.", Category = "ListItems")]
    public class RemoveIaCListItems : IaCCmdlet
    {
        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
        public string ListTitle { get; set; }

        /// <summary>
        /// Represents the start with ID to filter results based on ID specified
        /// </summary>
        [Parameter(Mandatory = false, ValueFromPipeline = true, Position = 1)]
        public int StartsWithId = 0;

        /// <summary>
        /// Represents the start with ID to filter results based on ID specified
        /// </summary>
        [Parameter(Mandatory = false, ValueFromPipeline = true, Position = 2)]
        public int EndsWithId = 0;

        /// <summary>
        /// Provides a xml caml query to execute in lue of running by IDs
        /// </summary>
        [Parameter(Mandatory = false, ValueFromPipeline = true, Position = 3)]
        public string OverrideCamlQuery { get; set; }

        /// <summary>
        /// Override the ViewFields to be returned in the Query
        /// </summary>
        [Parameter(Mandatory = false, ValueFromPipeline = true, Position = 4)]
        public string[] ViewFields { get; set; }

        /// <summary>
        /// Process the internals of the CmdLet
        /// </summary>
        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();

            // Load email notification list
            var listInSite = this.ClientContext.Web.Lists.GetByTitle(this.ListTitle);
            this.ClientContext.Load(listInSite);
            this.ClientContext.ExecuteQuery();

            // get site and query the list for pending requests
            var fieldNames = new List<string>() { "ID" };
            if (ViewFields != null && ViewFields.Length > 0)
            {
                fieldNames.AddRange(ViewFields);
            };

            var fieldsXml = CAML.ViewFields(fieldNames.Select(s => CAML.FieldRef(s)).ToArray());
            var camlWhereClause = string.Empty;
            var camlWhereConcat = false;

            if (!string.IsNullOrEmpty(this.OverrideCamlQuery))
            {
                camlWhereConcat = true;
                camlWhereClause = this.OverrideCamlQuery;
            }

            if (StartsWithId > 0)
            {
                var camlWhereSubClause = CAML.Geq(CAML.FieldValue("ID", FieldType.Number.ToString("f"), StartsWithId.ToString()));
                if (camlWhereConcat)
                {
                    // Wrap in an And Clause
                    camlWhereClause = CAML.And(camlWhereClause, camlWhereSubClause);
                }
                else
                {
                    camlWhereConcat = true;
                    camlWhereClause = camlWhereSubClause;
                }
            }

            if (EndsWithId > 0)
            {
                var camlWhereSubClause = CAML.Leq(CAML.FieldValue("ID", FieldType.Number.ToString("f"), EndsWithId.ToString()));
                if (camlWhereConcat)
                {
                    // Wrap in an And Clause
                    camlWhereClause = CAML.And(camlWhereClause, camlWhereSubClause);
                }
                else
                {
                    camlWhereConcat = true;
                    camlWhereClause = camlWhereSubClause;
                }
            }


            if (string.IsNullOrEmpty(camlWhereClause))
            {
                throw new Exception("Failed to construct a valid CAML Query.");
            }

            try
            {
                ListItemCollectionPosition ListItemCollectionPosition = null;
                var camlQuery = new CamlQuery()
                {
                    ViewXml = CAML.ViewQuery(ViewScope.RecursiveAll, CAML.Where(camlWhereClause), string.Empty, fieldsXml, 50)
                };

                var ids = new List<int>();
                while (true)
                {
                    camlQuery.ListItemCollectionPosition = ListItemCollectionPosition;
                    var spListItems = listInSite.GetItems(camlQuery);
                    this.ClientContext.Load(spListItems);
                    this.ClientContext.ExecuteQueryRetry();
                    ListItemCollectionPosition = spListItems.ListItemCollectionPosition;

                    foreach (var spListItem in spListItems)
                    {
                        var s = string.Empty;
                        LogWarning("ListItem [{0}] will be deleted.", spListItem.Id);
                        foreach (var fieldName in fieldNames)
                        {
                            s += string.Format("...[{0}]==[{1}]...", fieldName, spListItem[fieldName]);
                        }
                        LogVerbose("LISTITEM: {0}", s);
                        ids.Add(spListItem.Id);
                    }

                    if (ListItemCollectionPosition == null)
                    {
                        break;
                    }

                }

                // enumerate the ids to be deleted and process if -WhatIf (not passed)
                foreach (var id in ids)
                {
                    if (this.ShouldProcess(string.Format("ListItem [{0}] now being deleted.", id)))
                    {
                        var spListItem = listInSite.GetItemById(id);
                        spListItem.DeleteObject();
                        listInSite.Update();
                        this.ClientContext.ExecuteQueryRetry();
                    }
                }
            }
            catch (Exception ex)
            {
                LogError(ex, "Failed to query list {0} with message {1}", this.ListTitle, ex.Message);
            }
        }
    }
}
