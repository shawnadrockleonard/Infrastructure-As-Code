using InfrastructureAsCode.Core.Extensions;
using InfrastructureAsCode.Core.Models;
using InfrastructureAsCode.Powershell.CmdLets;
using InfrastructureAsCode.Powershell.PipeBinds;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Utilities;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Powershell.Commands.Lists
{
    /// <summary>
    /// Get List Items for the specifed list
    /// </summary>
    /// <remarks>
    /// Get-IaCListItems -Identity /Lists/Announcements
    /// Get-IaCListItems -Identity "Announcements"
    /// </remarks>
    [Cmdlet(VerbsCommon.Get, "IaCListItems", SupportsShouldProcess = true)]
    [CmdletHelp("Query the list and output the list items.", Category = "ListItems")]
    [OutputType(typeof(Collection<SPListItemDefinition>))]
    public class GetIaCListItems : IaCCmdlet
    {
        #region Parameters

        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0, HelpMessage = "The ID or Url of the list.")]
        public ListPipeBind Identity { get; set; }

        /// <summary>
        /// Optional caml entry for testing the query
        /// </summary>
        [Parameter(Mandatory = false)]
        public string CamlStatement { get; set; }

        /// <summary>
        /// Should we expand the list item to include its column data
        /// </summary>
        [Parameter(Mandatory = false)]
        public SwitchParameter Expand { get; set; }

        #endregion


        /// <summary>
        /// Execute the cmdlet
        /// </summary>
        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();



            Collection<SPListItemDefinition> results = new Collection<SPListItemDefinition>();

            var ctx = this.ClientContext;
            var _list = Identity.GetList(this.ClientContext.Web);
            ctx.Load(ctx.Web, wctx => wctx.Url, wctx => wctx.ServerRelativeUrl);
            ctx.Load(_list, lctx => lctx.ItemCount, lctx => lctx.EnableFolderCreation, lctx => lctx.RootFolder, lctx => lctx.RootFolder.ServerRelativeUrl, lctx => lctx.RootFolder.Folders);
            ctx.ExecuteQueryRetry();

            var webUri = new Uri(ctx.Web.Url);

            // Will query the list to determine the last item id in the list
            var lastItemId = _list.QueryLastItemId();
            LogVerbose("List with item count {0} has a last ID of {1}", _list.ItemCount, lastItemId);
            LogVerbose("List has folder creation = {0}", _list.EnableFolderCreation);


            var camlFields = new string[] { "Title", "ID", "Author", "Editor" };
            var camlViewFields = CAML.ViewFields(camlFields.Select(s => CAML.FieldRef(s)).ToArray());


            ListItemCollectionPosition itemPosition = null;
            var camlQueries = _list.SafeCamlClauseFromThreshold(2000, CamlStatement, 0, lastItemId);
            foreach (var camlAndValue in camlQueries)
            {
                itemPosition = null;
                var camlQuery = new CamlQuery()
                {
                    ViewXml = CAML.ViewQuery(
                        ViewScope.RecursiveAll,
                        CAML.Where(camlAndValue),
                        CAML.OrderBy(new OrderByField("ID")),
                        camlViewFields,
                        5)
                };

                LogWarning("CAML Query {0}", camlQuery.ViewXml);

                try
                {
                    while (true)
                    {
                        camlQuery.ListItemCollectionPosition = itemPosition;
                        var listItems = _list.GetItems(camlQuery);
                        _list.Context.Load(listItems, lti => lti.ListItemCollectionPosition);
                        _list.Context.ExecuteQueryRetry();
                        itemPosition = listItems.ListItemCollectionPosition;

                        foreach (var rbiItem in listItems)
                        {
                            var itemTitle = rbiItem.RetrieveListItemValue("Title");
                            LogVerbose("Title: {0}; Item ID: {1}", itemTitle, rbiItem.Id);

                            var newitem = new SPListItemDefinition()
                            {
                                Title = itemTitle,
                                Id = rbiItem.Id
                            };

                            if (Expand)
                            {
                                var author = rbiItem.RetrieveListItemUserValue("Author");
                                var editor = rbiItem.RetrieveListItemUserValue("Editor");
                                newitem.CreatedBy = new SPPrincipalUserDefinition()
                                {
                                    Email = author.ToUserEmailValue(),
                                    LoginName = author.ToUserValue(),
                                    Id = author.LookupId
                                };
                                newitem.ModifiedBy = new SPPrincipalUserDefinition()
                                {
                                    Email = editor.ToUserEmailValue(),
                                    LoginName = editor.ToUserValue(),
                                    Id = editor.LookupId
                                };

                                foreach (var rbiField in rbiItem.FieldValues)
                                {
                                    newitem.ColumnValues.Add(new SPListItemFieldDefinition()
                                    {
                                        FieldName = rbiField.Key,
                                        FieldValue = rbiField.Value
                                    });
                                }
                            }

                            results.Add(newitem);
                        }

                        if (itemPosition == null)
                        {
                            break;
                        }
                    }
                }
                catch (Exception ex)
                {
                    LogError(ex, "Failed to retrieve list item collection");
                }
            }

            WriteObject(results, true);
        }
    }
}

