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

namespace InfrastructureAsCode.Powershell.Commands.ListItems
{
    /// <summary>
    /// Query the specific list and delete items, if begin/end is specified filter the query
    /// </summary>
    [Cmdlet(VerbsCommon.Remove, "IaCListItems", SupportsShouldProcess = true)]
    public class RemoveIaCListItems : IaCCmdlet
    {
        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
        public string LibraryName { get; set; }

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

        [Parameter(Mandatory = false, ValueFromPipeline = true, Position = 4)]
        public object[] ViewFields { get; set; }

        /// <summary>
        /// Process the internals of the CmdLet
        /// </summary>
        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();

            //Load email notification list
            var listInSite = this.ClientContext.Web.Lists.GetByTitle(this.LibraryName);
            this.ClientContext.Load(listInSite);
            this.ClientContext.ExecuteQuery();

            // get ezforms site and query the list for pending requests
            var fieldNames = new List<string>() { "ID" };
            if (ViewFields.Length > 0)
            {
                fieldNames.AddRange(ViewFields.Select(s => s.ToString()));
            };

            var fieldsXml = fieldNames.Select(s => CAML.FieldRef(s)).ToArray();
            var camlWhereClause = string.Empty;
            var camlWhereConcat = false;

            if (!string.IsNullOrEmpty(this.OverrideCamlQuery))
            {
                camlWhereConcat = true;
                camlWhereClause = this.OverrideCamlQuery;
            }

            if (StartsWithId > 0)
            {
                var camlWhereSubClause = CAML.Gt(CAML.FieldValue("ID", "Number",  StartsWithId.ToString()));
                if (camlWhereConcat)
                {
                    // Wrap in an And Clause
                    camlWhereClause = CAML.And( camlWhereClause, camlWhereSubClause);
                }
                else
                {
                    camlWhereConcat = true;
                    camlWhereClause = camlWhereSubClause;
                }
            }

            if (EndsWithId > 0)
            {
                var camlWhereSubClause = CAML.Lt(CAML.FieldValue("ID", "Number", EndsWithId.ToString()));
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
                var camlQuery = new CamlQuery
                {
                    ViewXml = CAML.ViewQuery(ViewScope.RecursiveAll, CAML.Where(camlWhereClause), string.Empty, CAML.ViewFields(fieldsXml), 50)
                };

                var ids = new List<int>();
                while (true)
                {
                    camlQuery.ListItemCollectionPosition = ListItemCollectionPosition;
                    var spListItems = listInSite.GetItems(camlQuery);
                    this.ClientContext.Load(spListItems);
                    this.ClientContext.ExecuteQuery();
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

                if (this.ShouldProcess(string.Format("This will delete {0} list items.", ids.Count())))
                {
                    foreach (var id in ids)
                    {
                        LogWarning("ListItem [{0}] now being deleted.", id);
                        var spListItem = listInSite.GetItemById(id);
                        spListItem.DeleteObject();
                        listInSite.Update();
                        this.ClientContext.ExecuteQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                LogError(ex, "Failed to query list {0} with message {1}", this.LibraryName, ex.Message);
            }
        }
    }
}
