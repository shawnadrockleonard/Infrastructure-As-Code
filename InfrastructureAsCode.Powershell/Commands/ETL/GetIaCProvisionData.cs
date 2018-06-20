using InfrastructureAsCode.Core;
using InfrastructureAsCode.Core.Extensions;
using InfrastructureAsCode.Core.Models;
using InfrastructureAsCode.Powershell.CmdLets;
using InfrastructureAsCode.Powershell.PipeBinds;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;

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
            var contextWeb = this.ClientContext.Web;
            ctx.Load(contextWeb, ctxw => ctxw.ServerRelativeUrl, wctx => wctx.Url, ctxw => ctxw.Id);
            ctx.ExecuteQueryRetry();

            var _targetList = Identity.GetList(contextWeb, lctx => lctx.ItemCount, lctx => lctx.EnableFolderCreation, lctx => lctx.RootFolder, lctx => lctx.RootFolder.ServerRelativeUrl, lctx => lctx.RootFolder.Folders);
            
            
            var webUri = new Uri(contextWeb.Url);

            // Will query the list to determine the last item id in the list
            var lastItemId = _targetList.QueryLastItemId();
            logger.LogInformation("List with item count {0} has a last ID of {1}", _targetList.ItemCount, lastItemId);
            logger.LogInformation("List has folder creation = {0}", _targetList.EnableFolderCreation);

            // store the OOTB standard fields
            var ootbCamlFields = new List<string>() { "Title", "ID", "Author", "Editor", "Created", "Modified" };

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
            // wire up temporary array
            var sitelists = new List<SPListDefinition>()
            {
                ctx.GetListDefinition(_targetList, ExpandObjects, logger, skiptypes)
            };

            if (sitelists.Any())
            {
                var idx = 0;

                // lets add any list with NO lookups first
                var nolookups = sitelists.Where(sl => !sl.ListDependency.Any());
                nolookups.ToList().ForEach(nolookup =>
                {
                    logger.LogInformation("adding list {0}", nolookup.ListName);
                    nolookup.ProvisionOrder = idx++;
                    SiteComponents.Lists.Add(nolookup);
                    sitelists.Remove(nolookup);
                });

                // order with first in stack 
                var haslookups = sitelists.Where(sl => sl.ListDependency.Any()).OrderBy(ob => ob.ListDependency.Count()).ToList();
                while (haslookups.Count() > 0)
                {
                    var listPopped = haslookups.FirstOrDefault();
                    haslookups.Remove(listPopped);
                    logger.LogInformation("popping list {0}", listPopped.ListName);

                    if (listPopped.ListDependency.Any(listField => listPopped.ListName.Equals(listField, StringComparison.InvariantCultureIgnoreCase)
                                                          || listPopped.InternalName.Equals(listField, StringComparison.InvariantCultureIgnoreCase)))
                    {
                        // the list has a dependency on itself
                        logger.LogInformation("Adding list {0} with dependency on itself", listPopped.ListName);
                        listPopped.ProvisionOrder = idx++;
                        SiteComponents.Lists.Add(listPopped);
                    }
                    else if (listPopped.ListDependency.Any(listField =>
                            !SiteComponents.Lists.Any(sc => sc.ListName.Equals(listField, StringComparison.InvariantCultureIgnoreCase)
                                                         || sc.InternalName.Equals(listField, StringComparison.InvariantCultureIgnoreCase))))
                    {
                        // no list definition exists in the collection with the dependent lookup lists
                        logger.LogWarning("List {0} depends on {1} which do not exist current collection", listPopped.ListName, string.Join(",", listPopped.ListDependency));
                        foreach (var dependent in listPopped.ListDependency.Where(dlist => !haslookups.Any(hl => hl.ListName == dlist)))
                        {
                            var sitelist = contextWeb.GetListByTitle(dependent, lctx => lctx.Id, lctx => lctx.Title, lctx => lctx.RootFolder.ServerRelativeUrl);
                            var sitelistDefinition = ctx.GetListDefinition(contextWeb, sitelist, true, logger, skiptypes);
                            haslookups.Add(sitelistDefinition);
                        }

                        haslookups.Add(listPopped); // add back to collection
                        listPopped = null;
                    }
                    else
                    {
                        logger.LogInformation("Adding list {0} to collection", listPopped.ListName);
                        listPopped.ProvisionOrder = idx++;
                        SiteComponents.Lists.Add(listPopped);
                    }
                }
            }

            foreach (var listDefinition in SiteComponents.Lists.OrderBy(ob => ob.ProvisionOrder))
            {
                listDefinition.ListItems = new List<SPListItemDefinition>();

                var camlFields = new List<string>();
                var listTitle = listDefinition.ListName;
                var etlList = this.ClientContext.Web.GetListByTitle(listTitle, 
                    lctx => lctx.Id, lctx => lctx.ItemCount, lctx => lctx.EnableFolderCreation, lctx => lctx.RootFolder.ServerRelativeUrl);

                if (ExpandObjects)
                {
                    // list fields
                    var listFields = listDefinition.FieldDefinitions;
                    if (listFields != null && listFields.Any())
                    {
                        var filteredListFields = listFields.Where(lf => !skiptypes.Any(st => lf.FieldTypeKind == st)).ToList();
                        var notInCamlFields = listFields.Where(listField => !ootbCamlFields.Any(cf => cf.Equals(listField.InternalName, StringComparison.CurrentCultureIgnoreCase)) && !skiptypes.Any(st => listField.FieldTypeKind == st));
                        foreach (var listField in notInCamlFields)
                        {
                            logger.LogInformation("Processing list {0} field {1}", etlList.Title, listField.InternalName);
                            camlFields.Add(listField.InternalName);
                        }
                    }
                }

                camlFields.AddRange(ootbCamlFields);
                var camlViewFields = CAML.ViewFields(camlFields.Select(s => CAML.FieldRef(s)).ToArray());


                ListItemCollectionPosition itemPosition = null;
                var camlQueries = etlList.SafeCamlClauseFromThreshold(2000, CamlStatement, 0, lastItemId);
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
                            var listItems = etlList.GetItems(camlQuery);
                            etlList.Context.Load(listItems);
                            etlList.Context.ExecuteQueryRetry();
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
                
            }


            WriteObject(SiteComponents);
        }
    }
}

