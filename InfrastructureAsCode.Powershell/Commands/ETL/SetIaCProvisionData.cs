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

namespace InfrastructureAsCode.Powershell.Commands
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

        /// <summary>
        /// Represents the JSON file containing data
        /// </summary>
        [Parameter(Mandatory = true, HelpMessage = "Provide a full path to the JSON file containing data", Position = 1, ValueFromPipeline = true)]
        public string DataFilePath { get; set; }

        /// <summary>
        /// Specific list to be updated from the above action list
        /// </summary>
        [Parameter(Mandatory = true, Position = 2, ValueFromPipeline = true)]
        public ListPipeBind ListName { get; set; }


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

            if (!System.IO.File.Exists(this.DataFilePath))
            {
                var fileinfo = new System.IO.FileInfo(DataFilePath);
                throw new System.IO.FileNotFoundException("The data file was not found", fileinfo.Name);
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

            var listItemDefinitions = new List<SPListItemDefinition>();
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
                LogError(ex, "Failed to parse {0} Exception {1}", ProvisionerFilePath, ex.Message);
                return;
            }

            try
            {
                var dataFilePath = new System.IO.FileInfo(this.DataFilePath);
                var listItemJSON = System.IO.File.ReadAllText(dataFilePath.FullName);
                listItemDefinitions = JsonConvert.DeserializeObject<List<SPListItemDefinition>>(listItemJSON);
            }
            catch (Exception ex)
            {
                LogError(ex, "Failed to parse {0} Exception {1}", DataFilePath, ex.Message);
                return;
            }


            var etlList = ListName.GetList(this.ClientContext.Web,
                lctx => lctx.Id, lctx => lctx.RootFolder.ServerRelativeUrl, lctx => lctx.Title);
            LogVerbose("List {0}", etlList.Title);


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


            var customListDefinition = siteDefinition.Lists.FirstOrDefault(f => f.ListName.Equals(etlList.Title, StringComparison.InvariantCultureIgnoreCase)
                || f.InternalName.Equals(etlList.Title, StringComparison.InvariantCultureIgnoreCase));
            var customFields = customListDefinition.FieldDefinitions.Select(s => s.InternalName);


            foreach (var item in listItemDefinitions)
            {
                ListItemCreationInformation lici = new ListItemCreationInformation();
                ListItem newItem = etlList.AddItem(lici);
                newItem[ConstantsFields.Field_Title] = item.Title;
                newItem["SourceItemID"] = item.Id;
                newItem["DataMigrated"] = true;
                LogVerbose("Setting up Item {0} with Source ID {1}", item.Title, item.Id);

                var customColumns = item.ColumnValues.Where(cv => customFields.Any(cf => cf.Equals(cv.FieldName)));
                foreach (var spRefCol in customColumns)
                {
                    newItem[spRefCol.FieldName] = spRefCol.FieldValue;
                }
                newItem.Update();
                etlList.Context.ExecuteQueryRetry();
            }

        }
    }
}
