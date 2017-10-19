using InfrastructureAsCode.Powershell;
using InfrastructureAsCode.Powershell.CmdLets;
using InfrastructureAsCode.Core.Models;
using InfrastructureAsCode.Core.Extensions;
using InfrastructureAsCode.Powershell.Extensions;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using OfficeDevPnP.Core.Entities;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using InfrastructureAsCode.Powershell.PipeBinds;
using InfrastructureAsCode.Core.Constants;
using InfrastructureAsCode.Core.Reports;
using OfficeDevPnP.Core.Utilities;

namespace InfrastructureAsCode.Powershell.Commands.ETL
{
    /// <summary>
    /// The function cmdlet will accept the site provisioning template and a JSON file containing list data.  
    ///     It will bulk add the data or clobber the data
    /// </summary>
    [Cmdlet(VerbsCommon.Set, "IaCProvisionData", SupportsShouldProcess = true, ConfirmImpact = ConfirmImpact.High)]
    [CmdletHelp("Set list item data.", Category = "ETL")]
    public class SetIaCProvisionData : IaCCmdlet
    {
        #region Parameters

        /// <summary>
        /// Represents the directory path for any JSON files for serialization
        /// </summary>
        [Parameter(Mandatory = true, HelpMessage = "Provide a full path to the provisioner JSON file", Position = 0, ValueFromPipeline = true)]
        public string ProvisionerFilePath { get; set; }


        #endregion


        /// <summary>
        /// Validate parameters
        /// </summary>
        protected override void OnBeginInitialize()
        {
            if (!System.IO.File.Exists(this.ProvisionerFilePath))
            {
                var fileinfo = new System.IO.FileInfo(ProvisionerFilePath);
                throw new System.IO.FileNotFoundException("The provisioner file was not found", fileinfo.Name);
            }
        }

        /// <summary>
        /// Process the request
        /// </summary>
        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();

            // Initialize logging instance with Powershell logger
            ITraceLogger logger = new DefaultUsageLogger(LogVerbose, LogWarning, LogError);

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

            // Definition file to operate on
            SiteProvisionerModel siteDefinition = null;

            try
            {
                // Retreive JSON Provisioner file and deserialize it
                var filePath = new System.IO.FileInfo(this.ProvisionerFilePath);
                var filePathJSON = System.IO.File.ReadAllText(filePath.FullName);
                siteDefinition = JsonConvert.DeserializeObject<SiteProvisionerModel>(filePathJSON);
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Failed to parse {0} Exception {1}", ProvisionerFilePath, ex.Message);
                return;
            }


            foreach (var siteList in siteDefinition.Lists.OrderBy(ob => ob.ProvisionOrder))
            {
                var etlList = this.ClientContext.Web.GetListByTitle(siteList.ListName,
                    lctx => lctx.Id, lctx => lctx.RootFolder.ServerRelativeUrl, lctx => lctx.Title);
                logger.LogInformation("List {0}", etlList.Title);


                var dataMigratedFieldModel = new SPFieldDefinitionModel(FieldType.Boolean)
                {
                    InternalName = "DataMigrated",
                    Title = "DataMigrated",
                    AddToDefaultView = false,
                    AutoIndexed = true,
                    DefaultValue = "0",
                    FieldGuid = new Guid("9a353694-a5b4-4be4-a77a-c67bb071fbb5"),
                    FieldIndexed = true,
                    GroupName = "customcolumns"
                };
                var dataMigratedField = etlList.CreateListColumn(dataMigratedFieldModel, logger, null, null);

                var sourceFieldModel = new SPFieldDefinitionModel(FieldType.Integer)
                {
                    InternalName = "SourceItemID",
                    Title = "SourceItemID",
                    AddToDefaultView = false,
                    AutoIndexed = true,
                    FieldGuid = new Guid("e57a0936-5e8b-45d6-8c1c-d3e971c5b570"),
                    FieldIndexed = true,
                    GroupName = "customcolumns"
                };
                var sourceField = etlList.CreateListColumn(sourceFieldModel, logger, null, null);

                // Expectation is that list already exists in target location
                var customListDefinition = siteDefinition.Lists.FirstOrDefault(f => f.ListName.Equals(etlList.Title, StringComparison.InvariantCultureIgnoreCase)
                    || f.InternalName.Equals(etlList.Title, StringComparison.InvariantCultureIgnoreCase));

                // pull the internal names from list definition
                var customFields = customListDefinition.FieldDefinitions
                    .Where(lf => !skiptypes.Any(st => lf.FieldTypeKind == st))
                    .Select(s => s.InternalName);
                logger.LogWarning("Processing list {0} found {1} fields to be processed", etlList.Title, customFields.Count());

                // enumerate list items and add them to the list

                if (siteList.ListItems != null && siteList.ListItems.Any())
                {
                    foreach (var item in siteList.ListItems)
                    {
                        ListItemCreationInformation lici = new ListItemCreationInformation();
                        ListItem newItem = etlList.AddItem(lici);
                        newItem[ConstantsFields.Field_Title] = item.Title;
                        newItem["SourceItemID"] = item.Id;
                        newItem["DataMigrated"] = true;
                        logger.LogInformation("Setting up Item {0} with Source ID {1}", item.Title, item.Id);

                        var customColumns = item.ColumnValues.Where(cv => customFields.Any(cf => cf.Equals(cv.FieldName)));
                        foreach (var spRefCol in customColumns)
                        {
                            SPFieldDefinitionModel strParent = null;
                            var internalFieldName = spRefCol.FieldName;
                            var itemColumnValue = spRefCol.FieldValue;

                            if (IsLookup(siteList, internalFieldName, out strParent))
                            {
                                newItem[internalFieldName] = GetParentItemID(this.ClientContext, itemColumnValue, strParent, logger);
                            }
                            else if (IsUser(siteList, internalFieldName))
                            {
                                newItem[internalFieldName] = GetUserID(this.ClientContext, itemColumnValue, logger);
                            }
                            else
                            {
                                newItem[internalFieldName] = itemColumnValue;
                            }
                        }
                        newItem.Update();
                        etlList.Context.ExecuteQueryRetry();
                    }
                }
            }
        }

        /// <summary>
        /// Determines if the column is a lookup
        /// </summary>
        /// <param name="List"></param>
        /// <param name="ColumnName"></param>
        /// <param name="LookupListName"></param>
        /// <returns></returns>
        static bool IsLookup(SPListDefinition List, string ColumnName, out SPFieldDefinitionModel LookupField)
        {
            LookupField = null;
            foreach (var spc in List.FieldDefinitions)
            {
                if (spc.InternalName == ColumnName && spc.FieldTypeKind == FieldType.Lookup)
                {
                    LookupField = spc;
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Determines if the column is a lookup
        /// </summary>
        /// <param name="List"></param>
        /// <param name="ColumnName"></param>
        /// <returns></returns>
        static bool IsUser(SPListDefinition List, string ColumnName)
        {
            foreach (var spc in List.FieldDefinitions)
            {
                if (spc.InternalName == ColumnName && spc.FieldTypeKind == FieldType.User)
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Returns the parent item ID
        /// </summary>
        /// <param name="cContext"></param>
        /// <param name="ItemName"></param>
        /// <param name="logger">diagnostics logger</param>
        /// <returns></returns>
        static int GetUserID(ClientContext cContext, string ItemName, ITraceLogger logger)
        {
            int nReturn = -1;

            try
            {
                logger.LogInformation("Start GetUserID {0}", ItemName);
                Web wWeb = cContext.Web;
                var iUser = cContext.Web.EnsureUser(ItemName);
                cContext.Load(iUser);
                cContext.ExecuteQueryRetry();
                if (iUser != null)
                {
                    return iUser.Id;
                }
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Failed to find {0} in web {1}", ItemName, ex.Message);
            }

            return nReturn;
        }

        /// <summary>
        /// Returns the parent item ID
        /// </summary>
        /// <param name="cContext"></param>
        /// <param name="ItemName"></param>
        /// <param name="ParentListColumn"></param>
        /// <param name="logger">diagnostics logger</param>
        /// <returns></returns>
        static int GetParentItemID(ClientContext cContext, string ItemName, SPFieldDefinitionModel ParentListColumn, ITraceLogger logger)
        {
            int nReturn = -1;
            var parentListName = string.Empty;

            try
            {
                parentListName = ParentListColumn.Title;
                logger.LogInformation("Start GetParentItemID {0} for column {1}", parentListName, ItemName);

                Web wWeb = cContext.Web;
                
                var lParentList = cContext.Web.GetListByTitle(parentListName, lctx => lctx.Id, lctx => lctx.Title);
                var camlQuery = new CamlQuery()
                {
                    ViewXml = CAML.ViewQuery(
                        CAML.Where(
                            CAML.Eq(
                                CAML.FieldValue(ParentListColumn.LookupListFieldName, FieldType.Text.ToString("f"), ItemName))
                            ),
                        string.Empty,
                        100
                    )
                };

                ListItemCollectionPosition itemPosition = null;
                while (true)
                {
                    var collListItem = lParentList.GetItems(camlQuery);
                    lParentList.Context.Load(collListItem, lictx => lictx.ListItemCollectionPosition);
                    lParentList.Context.ExecuteQueryRetry();
                    itemPosition = collListItem.ListItemCollectionPosition;

                    foreach (ListItem oListItem in collListItem)
                    {
                        if (oListItem["Title"].ToString() == ItemName)
                        {
                            nReturn = oListItem.Id;
                            break;
                        }
                    }

                    if(itemPosition == null)
                    {
                        break;
                    }
                }

                logger.LogInformation("Complete GetParentItemID {0} for column {1}", parentListName, ItemName);
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Failed to query lookup value {0}", ex.Message);
            }

            return nReturn;
        }
    }
}
