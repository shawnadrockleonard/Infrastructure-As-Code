using IaC.Powershell;
using IaC.Powershell.CmdLets;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;

namespace IaC.Powershell.Commands.Lists
{
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

        [Parameter(Mandatory = false, ValueFromPipeline = true, Position = 4)]
        public string[] ViewFields { get; set; }

        /// <summary>
        /// Process the internals of the CmdLet
        /// </summary>
        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();

            //Load email notification list
            var listInSite = this.ClientContext.Web.Lists.GetByTitle(this.ListTitle);
            this.ClientContext.Load(listInSite);
            this.ClientContext.ExecuteQuery();

            // get ezforms site and query the list for pending requests
            var fieldNames = new List<string>() { "ID" };
            if (ViewFields != null && ViewFields.Length > 0)
            {
                fieldNames.AddRange(ViewFields);
            };

            var fieldsXml = string.Join(string.Empty, fieldNames.Select(s => string.Format("<FieldRef Name='{0}'/>", s)));
            var camlWhereClause = string.Empty;
            var camlWhereConcat = false;

            if (!string.IsNullOrEmpty(this.OverrideCamlQuery))
            {
                camlWhereConcat = true;
                camlWhereClause = this.OverrideCamlQuery;
            }

            if (StartsWithId > 0)
            {
                var camlWhereSubClause = string.Format("<Gt><FieldRef Name='ID' /><Value Type='Number'>{0}</Value></Gt>", StartsWithId);
                if (camlWhereConcat)
                {
                    // Wrap in an And Clause
                    camlWhereClause = string.Format("<And>{0}{1}</And>", camlWhereClause, camlWhereSubClause);
                }
                else
                {
                    camlWhereConcat = true;
                    camlWhereClause = camlWhereSubClause;
                }
            }

            if (EndsWithId > 0)
            {
                var camlWhereSubClause = string.Format("<Lt><FieldRef Name='ID' /><Value Type='Number'>{0}</Value></Lt>", EndsWithId);
                if (camlWhereConcat)
                {
                    // Wrap in an And Clause
                    camlWhereClause = string.Format("<And>{0}{1}</And>", camlWhereClause, camlWhereSubClause);
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
                var camlQuery = CamlQuery.CreateAllItemsQuery();
                camlQuery.ViewXml = string.Format("<View><Query><Where>{0}</Where>", camlWhereClause);
                camlQuery.ViewXml += string.Format("<ViewFields>{0}</ViewFields>", fieldsXml);
                camlQuery.ViewXml += "<RowLimit>50</RowLimit>";
                camlQuery.ViewXml += "</Query></View>";
                camlQuery.ListItemCollectionPosition = ListItemCollectionPosition;

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

                if (!this.DoNothing)
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
                LogError(ex, "Failed to query list {0} with message {1}", this.ListTitle, ex.Message);
            }
        }
    }
}
