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
using System.Xml.Linq;

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



            var results = new Collection<SPListItemDefinition>();

            var ctx = this.ClientContext;
            var _list = Identity.GetList(this.ClientContext.Web);
            ctx.Load(ctx.Web, wctx => wctx.Url, wctx => wctx.ServerRelativeUrl);
            ctx.Load(_list, lctx => lctx.ItemCount, lctx => lctx.EnableFolderCreation, lctx => lctx.RootFolder, lctx => lctx.RootFolder.ServerRelativeUrl, lctx => lctx.RootFolder.Folders);

            // load the list and web properties
            ctx.ExecuteQueryRetry();

            var webUri = new Uri(ctx.Web.Url);

            // Will query the list to determine the last item id in the list
            var lastItemId = _list.QueryLastItemId();
            LogVerbose("List with item count {0} has a last ID of {1}", _list.ItemCount, lastItemId);
            LogVerbose("List has folder creation = {0}", _list.EnableFolderCreation);


            var camlFields = new List<string>() { "Title", "ID", "Author", "Editor", "Created", "Modified" };



            if (Expand)
            {
                // SharePoint URI for XML parsing
                XNamespace ns = "http://schemas.microsoft.com/sharepoint/";

                // Skip these specific fields
                var skiptypes = new FieldType[]
                {
                    FieldType.Invalid,
                    FieldType.Computed,
                    FieldType.ContentTypeId,
                    FieldType.Invalid,
                    FieldType.WorkflowStatus,
                    FieldType.WorkflowEventType,
                    FieldType.Threading,
                    FieldType.ThreadIndex,
                    FieldType.Recurrence,
                    FieldType.PageSeparator,
                    FieldType.OutcomeChoice,
                    FieldType.CrossProjectLink,
                    FieldType.ModStat,
                    FieldType.Error,
                    FieldType.MaxItems
                };

                var fields = ctx.LoadQuery(_list.Fields
                        .Include(
                            v => v.AutoIndexed,
                            v => v.CanBeDeleted,
                            v => v.DefaultFormula,
                            v => v.DefaultValue,
                            v => v.Description,
                            v => v.EnforceUniqueValues,
                            v => v.FieldTypeKind,
                            v => v.Filterable,
                            v => v.Group,
                            v => v.Hidden,
                            v => v.Id,
                            v => v.InternalName,
                            v => v.Indexed,
                            v => v.JSLink,
                            v => v.NoCrawl,
                            v => v.ReadOnlyField,
                            v => v.Required,
                            v => v.Title,
                            v => v.SchemaXml));
                ClientContext.ExecuteQueryRetry();


                foreach (var listField in fields)
                {
                    // skip internal fields
                    if (skiptypes.Any(st => listField.FieldTypeKind == st))
                    {
                        continue;
                    }

                    try
                    {
                        var fieldXml = listField.SchemaXml;
                        if (!string.IsNullOrEmpty(fieldXml))
                        {
                            var xdoc = XDocument.Parse(fieldXml, LoadOptions.PreserveWhitespace);
                            var xField = xdoc.Element("Field");
                            var xSourceID = xField.Attribute("SourceID");
                            var xScope = xField.Element("Scope");
                            if (xSourceID != null
                                && xSourceID.Value.IndexOf(ns.NamespaceName, StringComparison.CurrentCultureIgnoreCase) < 0
                                && !camlFields.Any(cf => cf.Equals(listField.InternalName, StringComparison.CurrentCultureIgnoreCase)))
                            {
                                camlFields.Add(listField.InternalName);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Trace.TraceError("Failed to parse field {0} MSG:{1}", listField.InternalName, ex.Message);
                    }
                }
            }



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
                            var itemId = rbiItem.Id;
                            var itemTitle = rbiItem.RetrieveListItemValue("Title");
                            var author = rbiItem.RetrieveListItemUserValue("Author");
                            var editor = rbiItem.RetrieveListItemUserValue("Editor");
                            LogVerbose("Title: {0}; Item ID: {1}", itemTitle, itemId);

                            var newitem = new SPListItemDefinition()
                            {
                                Id = itemId,
                                Title = itemTitle,
                                Created = rbiItem.RetrieveListItemValue("Created").ToNullableDatetime(),
                                Modified = rbiItem.RetrieveListItemValue("Modified").ToNullableDatetime()
                            };

                            if (author != null)
                            {
                                newitem.CreatedBy = new SPPrincipalUserDefinition()
                                {
                                    Email = author.ToUserEmailValue(),
                                    LoginName = author.ToUserValue(),
                                    Id = author.LookupId
                                };
                            }
                            if (editor != null)
                            {
                                newitem.ModifiedBy = new SPPrincipalUserDefinition()
                                {
                                    Email = editor.ToUserEmailValue(),
                                    LoginName = editor.ToUserValue(),
                                    Id = editor.LookupId
                                };
                            }


                            if (Expand)
                            {

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

