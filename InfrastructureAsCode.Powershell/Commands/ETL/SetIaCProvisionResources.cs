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
using InfrastructureAsCode.Core.Reports;

namespace InfrastructureAsCode.Powershell.Commands
{
    /// <summary>
    /// The function cmdlet will upgrade the site specified in the connection to the latest configuration changes
    /// </summary>
    [Cmdlet(VerbsCommon.Set, "IaCProvisionResources", SupportsShouldProcess = true)]
    [CmdletHelp("Set site definition components based on JSON file.", Category = "ETL")]
    public class SetIaCProvisionResources : IaCCmdlet
    {
        /// <summary>
        /// Represents the directory path for any JSON files for serialization
        /// </summary>
        [Parameter(Mandatory = true, HelpMessage = "Provide a full path to the provisioner JSON file", Position = 0, ValueFromPipeline = true)]
        public string ProvisionerFilePath { get; set; }

        /// <summary>
        /// Specific action list
        /// </summary>
        [ValidateNotNullOrEmpty()]
        [ValidateSet(new string[] { "ALL", "Groups", "ContentTypes", "Fields", "Lists", "Views", "ListData" })]
        [Parameter(Mandatory = false, ParameterSetName = "ActionDependency")]
        public string ActionSet { get; set; }

        /// <summary>
        /// Specific list to be updated from the above action list
        /// </summary>
        [Parameter(Mandatory = false, ParameterSetName = "ActionDependency")]
        public string SpecificListName { get; set; }

        /// <summary>
        /// Specific view to be updated from the above action list
        /// </summary>
        [Parameter(Mandatory = false, ParameterSetName = "ActionDependency")]
        public string SpecificViewName { get; set; }

        /// <summary>
        /// Specific view to be updated from the above action list
        /// </summary>
        [Parameter(Mandatory = false)]
        public SwitchParameter ValidateJson { get; set; }


        /// <summary>
        /// Holds the SharePoint groups in the site or created in the site
        /// </summary>
        private List<SPGroupDefinitionModel> siteGroups { get; set; }

        /// <summary>
        /// Holds the [Site] columns
        /// </summary>
        private List<SPFieldDefinitionModel> siteColumns { get; set; }

        /// <summary>
        /// Holds the [List] columns
        /// </summary>
        private List<SPFieldDefinitionModel> listColumns { get; set; }

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

            siteGroups = new List<SPGroupDefinitionModel>();
            siteColumns = new List<SPFieldDefinitionModel>();
            listColumns = new List<SPFieldDefinitionModel>();
        }

        /// <summary>
        /// Process the request
        /// </summary>
        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();

            // Initialize logging instance with Powershell logger
            ITraceLogger logger = new DefaultUsageLogger(LogVerbose, LogWarning, LogError);

            // SharePoint URI for XML parsing
            XNamespace ns = "http://schemas.microsoft.com/sharepoint/";

            // Retreive JSON Provisioner file and deserialize it
            var filePath = new System.IO.FileInfo(this.ProvisionerFilePath);
            SiteProvisionerModel siteDefinition = null;

            try
            {
                siteDefinition = JsonConvert.DeserializeObject<SiteProvisionerModel>(System.IO.File.ReadAllText(filePath.FullName));
            }
            catch (Exception ex)
            {
                LogError(ex, "Failed to parse {0} Exception {1}", filePath.Name, ex.Message);
                return;
            }


            if (ValidateJson)
            {
                LogVerbose("The file {0} is valid.", filePath.Name);
                return;
            }

            var provisionerChoices = siteDefinition.FieldChoices;

            // Load the Context
            var contextWeb = this.ClientContext.Web;
            this.ClientContext.Load(contextWeb,
                ctxw => ctxw.ServerRelativeUrl,
                ctxw => ctxw.Id,
                ctxw => ctxw.Fields.Include(inc => inc.InternalName, inc => inc.JSLink, inc => inc.Title, inc => inc.Id));

            // All Site Columns
            var siteFields = this.ClientContext.LoadQuery(ClientContext.Web.AvailableFields
                .Include(inc => inc.InternalName, inc => inc.JSLink, inc => inc.Title, inc => inc.Id));

            // pull Site Groups
            var groupQuery = this.ClientContext.LoadQuery(contextWeb.SiteGroups
                .Include(group => group.Id,
                        group => group.Title,
                        group => group.Description,
                        group => group.AllowRequestToJoinLeave,
                        group => group.AllowMembersEditMembership,
                        group => group.AutoAcceptRequestToJoinLeave,
                        group => group.OnlyAllowMembersViewMembership,
                        group => group.RequestToJoinLeaveEmailSetting));

            var contentTypes = this.ClientContext.LoadQuery(contextWeb.ContentTypes
                .Include(
                        ict => ict.Id,
                        ict => ict.Group,
                        ict => ict.Description,
                        ict => ict.Name,
                        ict => ict.Hidden,
                        ict => ict.JSLink,
                        ict => ict.FieldLinks,
                        ict => ict.Fields));


            var collists = contextWeb.Lists;
            var lists = this.ClientContext.LoadQuery(collists
                .Include(
                    linc => linc.Title,
                    linc => linc.Id,
                    linc => linc.Description,
                    linc => linc.Hidden,
                    linc => linc.OnQuickLaunch,
                    linc => linc.BaseTemplate,
                    linc => linc.ContentTypesEnabled,
                    linc => linc.AllowContentTypes,
                    linc => linc.EnableFolderCreation,
                    linc => linc.IsApplicationList,
                    linc => linc.IsCatalog,
                    linc => linc.IsSiteAssetsLibrary,
                    linc => linc.IsPrivate,
                    linc => linc.IsSystemList,
                    linc => linc.Views,
                    linc => linc.Fields
                    .Include(
                        lft => lft.Id,
                        lft => lft.AutoIndexed,
                        lft => lft.CanBeDeleted,
                        lft => lft.DefaultFormula,
                        lft => lft.DefaultValue,
                        lft => lft.Group,
                        lft => lft.Description,
                        lft => lft.EnforceUniqueValues,
                        lft => lft.FieldTypeKind,
                        lft => lft.Filterable,
                        lft => lft.Hidden,
                        lft => lft.Indexed,
                        lft => lft.InternalName,
                        lft => lft.JSLink,
                        lft => lft.NoCrawl,
                        lft => lft.ReadOnlyField,
                        lft => lft.Required,
                        lft => lft.SchemaXml,
                        lft => lft.Scope,
                        lft => lft.Title
                        ),
                    linc => linc.ContentTypes
                    .Include(
                        ict => ict.Id,
                        ict => ict.Group,
                        ict => ict.Description,
                        ict => ict.Name,
                        ict => ict.Hidden,
                        ict => ict.JSLink,
                        ict => ict.FieldLinks,
                        ict => ict.Fields)).Where(w => !w.IsSystemList && !w.IsSiteAssetsLibrary));
            this.ClientContext.ExecuteQueryRetry();



            /// Provision Site Groups
            /// Provision Site Columns
            /// Provision Site Content Type
            /// Provision List
            /// Create List Columns
            /// Associate Content Type to List
            /// Provision List Views
            /// 



            // creates groups
            contextWeb.Context.Load(contextWeb, hw => hw.CurrentUser);
            contextWeb.Context.ExecuteQueryRetry();


            if (string.IsNullOrEmpty(ActionSet) || ActionSet == "ALL" || ActionSet == "Groups")
            {
                if (siteDefinition.Groups != null && siteDefinition.Groups.Any())
                {
                    LogVerbose("Group collection will be provisioned for {0} groups", siteDefinition.Groups.Count());
                    foreach (var groupDef in siteDefinition.Groups)
                    {

                        if (groupQuery.Any(g => g.Title.Equals(groupDef.Title, StringComparison.CurrentCultureIgnoreCase)))
                        {
                            var group = groupQuery.FirstOrDefault(g => g.Title.Equals(groupDef.Title, StringComparison.CurrentCultureIgnoreCase));
                            siteGroups.Add(new SPGroupDefinitionModel() { Id = group.Id, Title = group.Title });
                        }
                        else
                        {

                            var newgroup = contextWeb.GetOrCreateSiteGroups(groupDef);
                            siteGroups.Add(new SPGroupDefinitionModel() { Id = newgroup.Id, Title = newgroup.Title });
                        }
                    }
                }
            }

            // provision columns
            // Site Columns
            if (string.IsNullOrEmpty(ActionSet) || ActionSet == "ALL" || ActionSet == "Fields")
            {
                if (siteDefinition.FieldDefinitions != null && siteDefinition.FieldDefinitions.Any())
                {
                    LogVerbose("Field definitions will be provisioned for {0} fields", siteDefinition.FieldDefinitions.Count());
                    foreach (var fieldDef in siteDefinition.FieldDefinitions)
                    {
                        var column = contextWeb.CreateColumn(fieldDef, logger, siteGroups, siteDefinition.FieldChoices);
                        if (column == null)
                        {
                            LogWarning("Failed to create column {0}.", fieldDef.InternalName);
                        }
                        else
                        {
                            siteColumns.Add(new SPFieldDefinitionModel()
                            {
                                InternalName = column.InternalName,
                                Title = column.Title,
                                FieldGuid = column.Id
                            });
                        }
                    }
                }
            }


            // provision content types
            if (string.IsNullOrEmpty(ActionSet) || ActionSet == "ALL" || ActionSet == "ContentTypes")
            {
                if (siteDefinition.ContentTypes != null && siteDefinition.ContentTypes.Any())
                {
                    LogVerbose("Content types will be provisioned for {0} ctypes", siteDefinition.ContentTypes.Count());
                    foreach (var contentDef in siteDefinition.ContentTypes)
                    {
                        var contentTypeName = contentDef.Name;
                        var contentTypeId = contentDef.ContentTypeId;

                        if (!contextWeb.ContentTypeExistsByName(contentTypeName)
                            && !contextWeb.ContentTypeExistsById(contentTypeId))
                        {
                            LogVerbose("Provisioning content type {0}", contentTypeName);
                            contextWeb.CreateContentType(contentTypeName, contentTypeId, (string.IsNullOrEmpty(contentDef.ContentTypeGroup) ? "CustomColumn" : contentDef.ContentTypeGroup));
                        }

                        var provisionedContentType = contextWeb.GetContentTypeByName(contentTypeName, true);
                        if (provisionedContentType != null)
                        {
                            LogVerbose("Found content type {0} and is read only {1}", contentTypeName, provisionedContentType.ReadOnly);
                            if (!provisionedContentType.ReadOnly)
                            {
                                foreach (var fieldDef in contentDef.FieldLinks)
                                {
                                    // Check if FieldLink exists in the Content Type
                                    if (!contextWeb.FieldExistsByNameInContentType(contentTypeName, fieldDef.Name))
                                    {
                                        // Check if FieldLInk column is in the collection of FieldDefinition
                                        var siteColumn = siteDefinition.FieldDefinitions.FirstOrDefault(f => f.InternalName == fieldDef.Name);
                                        if (siteColumn != null && !contextWeb.FieldExistsByNameInContentType(contentTypeName, siteColumn.DisplayNameMasked))
                                        {
                                            var column = this.siteColumns.FirstOrDefault(f => f.InternalName == fieldDef.Name);
                                            if (column == null)
                                            {
                                                LogWarning("Column {0} was not added to the collection", fieldDef.Name);
                                            }
                                            else
                                            {
                                                contextWeb.AddFieldToContentTypeByName(contentTypeName, column.FieldGuid, siteColumn.Required);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }


            // provision lists
            if (string.IsNullOrEmpty(ActionSet) || ActionSet == "ALL" || ActionSet == "Lists" || ActionSet == "Views" || ActionSet == "ListData")
            {
                var listtoprocess = siteDefinition.Lists
                    .Where(w => (string.IsNullOrEmpty(SpecificListName) || (!String.IsNullOrEmpty(SpecificListName) && w.ListName.Equals(SpecificListName, StringComparison.InvariantCultureIgnoreCase))))
                    .OrderBy(w => w.ProvisionOrder)
                    .ToList();

                foreach (var listDef in listtoprocess)
                {
                    // Content Type
                    var listName = listDef.ListName;
                    var listDescription = listDef.ListDescription;


                    // provision the list definition
                    var siteList = contextWeb.CreateListFromDefinition(listDef, provisionerChoices);

                    if (string.IsNullOrEmpty(ActionSet) || ActionSet == "ALL" || ActionSet == "Lists")
                    {
                        if (listDef.ContentTypeEnabled && listDef.HasContentTypes)
                        {
                            if (listDef.ContentTypes != null && listDef.ContentTypes.Any())
                            {
                                LogVerbose("List {0} => Content types will be provisioned for {1} ctypes", listDef.ListName, listDef.ContentTypes.Count());
                                foreach (var contentDef in listDef.ContentTypes)
                                {
                                    var contentTypeName = contentDef.Name;
                                    ContentType accessContentType = null;

                                    if (!contextWeb.ContentTypeExistsByName(listName, contentTypeName))
                                    {
                                        if (siteDefinition.ContentTypes != null && siteDefinition.ContentTypes.Any(ct => ct.Name == contentTypeName))
                                        {
                                            contextWeb.AddContentTypeToListByName(listName, contentTypeName, true);
                                            accessContentType = siteList.GetContentTypeByName(contentTypeName);
                                        }
                                        else
                                        {
                                            var ctypeInfo = contentDef.ToCreationObject();
                                            accessContentType = siteList.ContentTypes.Add(ctypeInfo);
                                            siteList.Update();
                                            siteList.Context.Load(accessContentType, tycp => tycp.Id, tycp => tycp.Name);
                                            siteList.Context.ExecuteQueryRetry();
                                        }

                                        if (contentDef.DefaultContentType)
                                        {
                                            siteList.SetDefaultContentTypeToList(accessContentType);
                                        }
                                    }
                                }
                            }
                        }

                        if (listDef.ListName == "CollectionSiteTypesLK")
                        {

                        }

                        // Existing columns
                        var internalNamesForList = listDef.FieldDefinitions.Select(s => s.InternalName).ToArray();
                        var internalNamesFoundInList = new List<string>();
                        try
                        {
                            var existingListColumns = siteList.GetFields(internalNamesForList);
                            foreach (var column in existingListColumns)
                            {
                                listColumns.Add(new SPFieldDefinitionModel()
                                {
                                    InternalName = column.InternalName,
                                    Title = column.Title,
                                    FieldGuid = column.Id
                                });
                                internalNamesFoundInList.Add(column.InternalName);
                            }
                        }
                        catch(Exception ex)
                        {
                            LogError(ex, "List {0} => failed to query columns by internal names {1}", listDef.ListName, ex.Message);
                        }

                        // List Columns
                        var nonExistingListColumns = listDef.FieldDefinitions.Where(fd => !internalNamesFoundInList.Any(inf => fd.InternalName.Equals(inf, StringComparison.InvariantCultureIgnoreCase)));
                        foreach (var fieldDef in nonExistingListColumns)
                        {
                            if (fieldDef.FromBaseType == true && fieldDef.SourceID.IndexOf(ns.NamespaceName, StringComparison.CurrentCultureIgnoreCase) > -1)
                            {
                                // OOTB Column
                                var hostsitecolumn = siteFields.FirstOrDefault(fd => fd.InternalName == fieldDef.InternalName);
                                if (hostsitecolumn != null && !siteList.FieldExistsByName(hostsitecolumn.InternalName))
                                {
                                    var column = siteList.Fields.Add(hostsitecolumn);
                                    siteList.Update();
                                    siteList.Context.Load(column, cctx => cctx.Id, cctx => cctx.InternalName);
                                    siteList.Context.ExecuteQueryRetry();
                                }

                                var sourceListColumns = siteList.GetFields(fieldDef.InternalName);
                                foreach (var column in sourceListColumns)
                                {
                                    listColumns.Add(new SPFieldDefinitionModel()
                                    {
                                        InternalName = column.InternalName,
                                        Title = column.Title,
                                        FieldGuid = column.Id
                                    });
                                }
                            }
                            else if (fieldDef.FieldTypeKind == FieldType.Invalid
                                && fieldDef.FieldTypeKindText.IndexOf("TaxonomyFieldType", StringComparison.InvariantCultureIgnoreCase) > -1)
                            {
                                // Taxonomy Column
                                var hostsitecolumn = siteFields.FirstOrDefault(fd => fd.InternalName == fieldDef.InternalName);
                                if (hostsitecolumn != null && !siteList.FieldExistsByName(hostsitecolumn.InternalName))
                                {
                                    var column = siteList.Fields.Add(hostsitecolumn);
                                    siteList.Update();
                                    siteList.Context.Load(column, cctx => cctx.Id, cctx => cctx.InternalName);
                                    siteList.Context.ExecuteQueryRetry();
                                }

                                var sourceListColumns = siteList.GetFields(fieldDef.InternalName);
                                foreach (var column in sourceListColumns)
                                {
                                    listColumns.Add(new SPFieldDefinitionModel()
                                    {
                                        InternalName = column.InternalName,
                                        Title = column.Title,
                                        FieldGuid = column.Id
                                    });
                                }
                            }
                            else
                            {
                                var column = siteList.CreateListColumn(fieldDef, logger, siteGroups, provisionerChoices);
                                if (column == null)
                                {
                                    LogWarning("Failed to create column {0}.", fieldDef.InternalName);
                                }
                                else
                                {
                                    listColumns.Add(new SPFieldDefinitionModel()
                                    {
                                        InternalName = column.InternalName,
                                        Title = column.Title,
                                        FieldGuid = column.Id
                                    });
                                }
                            }
                        }

                        // Where content types are enabled
                        // Add the provisioned site columns or list columns to the content type
                        if (listDef.ContentTypeEnabled && listDef.HasContentTypes)
                        {
                            foreach (var contentDef in listDef.ContentTypes)
                            {
                                var contentTypeName = contentDef.Name;
                                var accessContentTypes = siteList.ContentTypes;
                                IEnumerable<ContentType> allContentTypes = contextWeb.Context.LoadQuery(accessContentTypes.Where(f => f.Name == contentTypeName).Include(tcyp => tcyp.Id, tcyp => tcyp.Name));
                                contextWeb.Context.ExecuteQueryRetry();

                                if (allContentTypes != null)
                                {
                                    var accessContentType = allContentTypes.FirstOrDefault();
                                    foreach (var fieldInternalName in contentDef.FieldLinks)
                                    {
                                        var column = listColumns.FirstOrDefault(f => f.InternalName == fieldInternalName.Name);
                                        if (column == null)
                                        {
                                            LogWarning("List {0} => Failed to associate field link {1}.", listDef.ListName, fieldInternalName.Name);
                                            continue;
                                        }

                                        var fieldLinks = accessContentType.FieldLinks;
                                        contextWeb.Context.Load(fieldLinks, cf => cf.Include(inc => inc.Id, inc => inc.Name));
                                        contextWeb.Context.ExecuteQueryRetry();

                                        var convertedInternalName = column.DisplayNameMasked;
                                        if (!fieldLinks.Any(a => a.Name == column.InternalName || a.Name == convertedInternalName))
                                        {
                                            LogVerbose("List {0} => Content Type {1} Adding Field {2}", listDef.ListName, contentTypeName, column.InternalName);
                                            var siteColumn = siteList.GetFieldById<Field>(column.FieldGuid);
                                            contextWeb.Context.ExecuteQueryRetry();

                                            var flink = new FieldLinkCreationInformation();
                                            flink.Field = siteColumn;
                                            var flinkstub = accessContentType.FieldLinks.Add(flink);
                                            //if(fieldDef.Required) flinkstub.Required = fieldDef.Required;
                                            accessContentType.Update(false);
                                            contextWeb.Context.ExecuteQueryRetry();
                                        }
                                    }
                                }
                            }
                        }
                    }

                    // Views
                    if (string.IsNullOrEmpty(ActionSet) || ActionSet == "ALL" || ActionSet == "Lists" || ActionSet == "Views")
                    {
                        if (listDef.Views != null && listDef.Views
                            .Any(w => (string.IsNullOrEmpty(SpecificViewName) || (!String.IsNullOrEmpty(SpecificViewName) && w.Title.Equals(SpecificViewName, StringComparison.InvariantCultureIgnoreCase)))))
                        {
                            ViewCollection views = siteList.Views;
                            contextWeb.Context.Load(views, f => f.Include(inc => inc.Id, inc => inc.Hidden, inc => inc.Title, inc => inc.DefaultView));
                            contextWeb.Context.ExecuteQueryRetry();

                            foreach (var modelView in listDef.Views
                                .Where(w => (string.IsNullOrEmpty(SpecificViewName) || (!String.IsNullOrEmpty(SpecificViewName) && w.Title.Equals(SpecificViewName, StringComparison.InvariantCultureIgnoreCase)))))
                            {
                                try
                                {
                                    var updatecaml = false;
                                    View view = null;
                                    if (views.Any(v => v.Title.Equals(modelView.Title, StringComparison.CurrentCultureIgnoreCase)))
                                    {
                                        LogVerbose("List {0} => View {1} found in list", listName, modelView.Title);
                                        view = views.FirstOrDefault(v => v.Title.Equals(modelView.Title, StringComparison.CurrentCultureIgnoreCase));
                                        updatecaml = true;
                                    }
                                    else
                                    {
                                        LogVerbose("List {0} => Creating View {0} in list", listName, modelView.Title);
                                        view = siteList.CreateView(modelView.CalculatedInternalName, modelView.ViewCamlType, modelView.FieldRefName.ToArray(), modelView.RowLimit, modelView.DefaultView, modelView.ViewQuery, modelView.PersonalView, modelView.Paged);
                                    }

                                    // grab the view properties from the object
                                    view.EnsureProperties(
                                        mview => mview.Title,
                                        mview => mview.Scope,
                                        mview => mview.AggregationsStatus,
                                        mview => mview.Aggregations,
                                        mview => mview.DefaultView,
                                        mview => mview.Hidden,
                                        mview => mview.Toolbar,
                                        mview => mview.JSLink,
                                        mview => mview.ViewFields,
                                        vctx => vctx.ViewQuery
                                        );


                                    if (modelView.FieldRefName != null && modelView.FieldRefName.Any())
                                    {
                                        var currentFields = view.ViewFields;
                                        currentFields.RemoveAll();
                                        modelView.FieldRefName.ToList().ForEach(vField =>
                                        {
                                            currentFields.Add(vField.Trim());
                                        });
                                    }

                                    if (!string.IsNullOrEmpty(modelView.Aggregations))
                                    {
                                        view.Aggregations = modelView.Aggregations;
                                        view.AggregationsStatus = modelView.AggregationsStatus;
                                    }

                                    if (modelView.Hidden.HasValue && modelView.Hidden == true)
                                    {
                                        view.Hidden = modelView.Hidden.Value;
                                    }

                                    if (modelView.ToolBarType.HasValue)
                                    {
                                        view.Toolbar = string.Format("<Toolbar Type=\"{0}\"/>", modelView.ToolBarType.ToString());
                                    }

                                    if (updatecaml)
                                    {
                                        view.DefaultView = modelView.DefaultView;
                                        view.RowLimit = modelView.RowLimit;
                                        view.ViewQuery = modelView.ViewQuery;
                                    }

                                    if (modelView.HasJsLink && modelView.JsLink.IndexOf("clienttemplates.js") == -1)
                                    {
                                        view.JSLink = modelView.JsLink;
                                    }

                                    view.Scope = modelView.Scope;
                                    view.Title = modelView.Title;
                                    view.Update();
                                    contextWeb.Context.Load(view, v => v.Title, v => v.Id, v => v.ServerRelativeUrl);
                                    contextWeb.Context.ExecuteQueryRetry();
                                }
                                catch (Exception ex)
                                {
                                    LogError(ex, "List {0} => Failed to create view {1} with XML:{2}", listDef.ListName, modelView.Title, modelView.ViewQuery);
                                }
                            }
                        }
                    }

                    // List Data if provided in the JSON file, lets add it to the list
                    if (string.IsNullOrEmpty(ActionSet) || ActionSet == "ALL" || ActionSet == "ListData")
                    {
                        if (listDef.ListItems != null && listDef.ListItems.Any())
                        {
                            foreach (var listItem in listDef.ListItems)
                            {
                                // Process the record into the notification email list
                                var itemCreateInfo = new ListItemCreationInformation();

                                var newSPListItem = siteList.AddItem(itemCreateInfo);
                                newSPListItem["Title"] = listItem.Title;
                                foreach (var listItemData in listItem.ColumnValues)
                                {
                                    newSPListItem[listItemData.FieldName] = listItemData.FieldValue;
                                }

                                newSPListItem.Update();
                                contextWeb.Context.ExecuteQueryRetry();
                            }
                        }
                    }
                }
            }
        }
    }
}
