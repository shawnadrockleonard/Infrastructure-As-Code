using InfrastructureAsCode.Core.Constants;
using InfrastructureAsCode.Core.Extensions;
using InfrastructureAsCode.Core.Models;
using InfrastructureAsCode.Core.Reports;
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

namespace InfrastructureAsCode.Powershell.Commands.ETL
{
    /// <summary>
    /// Get List Items for the specifed list
    /// </summary>
    /// <remarks>
    /// Get-IaCListItems -Identity /Lists/Announcements
    /// Get-IaCListItems -Identity "Announcements"
    /// </remarks>
    [Cmdlet(VerbsCommon.Get, "IaCProvisionData", SupportsShouldProcess = true)]
    [CmdletHelp("Query the list and output the list items.", Category = "ETL")]
    [OutputType(typeof(SiteProvisionerModel))]
    public class GetIaCProvisionData : IaCCmdlet
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
        public SwitchParameter ExpandObjects { get; set; }

        #endregion


        /// <summary>
        /// Execute the cmdlet
        /// </summary>
        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();


            // Initialize logging instance with Powershell logger
            ITraceLogger logger = new DefaultUsageLogger(LogVerbose, LogWarning, LogError);


            // Construct the model
            var SiteComponents = new SiteProvisionerModel()
            {
                Lists = new List<SPListDefinition>()
            };

            var ctx = this.ClientContext;
            var _list = Identity.GetList(this.ClientContext.Web, lctx => lctx.ItemCount, lctx => lctx.EnableFolderCreation, lctx => lctx.RootFolder, lctx => lctx.RootFolder.ServerRelativeUrl, lctx => lctx.RootFolder.Folders);

            ctx.Load(ctx.Web, wctx => wctx.Url, wctx => wctx.ServerRelativeUrl);
            ctx.ExecuteQueryRetry();

            var webUri = new Uri(ctx.Web.Url);

            // Will query the list to determine the last item id in the list
            var lastItemId = _list.QueryLastItemId();
            logger.LogInformation("List with item count {0} has a last ID of {1}", _list.ItemCount, lastItemId);
            logger.LogInformation("List has folder creation = {0}", _list.EnableFolderCreation);


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
                FieldType.MaxItems,
                FieldType.Attachments
            };

            // pull a small portion of the list 
            var listDefinition = this.ClientContext.GetListDefinition(_list, ExpandObjects, logger, skiptypes);
            listDefinition.ListItems = new List<SPListItemDefinition>();


            var ootbCamlFields = new List<string>() { "Title", "ID", "Author", "Editor", "Created", "Modified" };
            var camlFields = new List<string>(ootbCamlFields);


            if (ExpandObjects)
            {
                // list fields
                var listFields = listDefinition.FieldDefinitions;
                if (listFields != null && listFields.Any())
                {
                    var filteredListFields = listFields.Where(lf => !skiptypes.Any(st => lf.FieldTypeKind == st)).ToList();
                    var notInCamlFields = listFields.Where(listField => !ootbCamlFields.Any(cf => cf.Equals(listField.InternalName, StringComparison.CurrentCultureIgnoreCase)));
                    foreach (var listField in notInCamlFields)
                    {
                        logger.LogInformation("Processing list {0} field {1}", _list.Title, listField.InternalName);
                        camlFields.Add(listField.InternalName);
                    }
                }
            }



            var camlViewFields = CAML.ViewFields(camlFields.Select(s => CAML.FieldRef(s)).ToArray());


            ListItemCollectionPosition itemPosition = null;
            var camlQueries = _list.SafeCamlClauseFromThreshold(2000, CamlStatement, 0, lastItemId);
            foreach (var camlAndValue in camlQueries)
            {
                itemPosition = null;
                var camlWhereClause = (string.IsNullOrEmpty(camlAndValue) ? string.Empty : CAML.Where(camlAndValue));
                var camlQuery = new CamlQuery()
                {
                    ViewXml = CAML.ViewQuery(
                        ViewScope.RecursiveAll,
                        camlWhereClause,
                        CAML.OrderBy(new OrderByField("ID")),
                        camlViewFields,
                        500)
                };

                try
                {
                    while (true)
                    {
                        logger.LogWarning("CAML Query {0} at position {1}", camlWhereClause, (itemPosition == null ? string.Empty : itemPosition.PagingInfo));
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
                            logger.LogInformation("Title: {0}; Item ID: {1}", itemTitle, itemId);

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


                            if (ExpandObjects)
                            {
                                var fieldValuesToWrite = rbiItem.FieldValues.Where(rfield => !ootbCamlFields.Any(oc => rfield.Key.Equals(oc, StringComparison.CurrentCultureIgnoreCase)));
                                foreach (var rbiField in fieldValuesToWrite)
                                {
                                    newitem.ColumnValues.Add(new SPListItemFieldDefinition()
                                    {
                                        FieldName = rbiField.Key,
                                        FieldValue = rbiField.Value
                                    });
                                }
                            }

                            listDefinition.ListItems.Add(newitem);
                        }

                        if (itemPosition == null)
                        {
                            break;
                        }
                    }
                }
                catch (Exception ex)
                {
                    logger.LogError(ex, "Failed to retrieve list item collection");
                }
            }


            SiteComponents.Lists.Add(listDefinition);



            WriteObject(SiteComponents);
        }
    }
}

