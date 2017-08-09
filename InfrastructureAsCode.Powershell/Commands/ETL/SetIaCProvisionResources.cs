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

namespace InfrastructureAsCode.Powershell.Commands
{
    /// <summary>
    /// The function cmdlet will upgrade the EzForms site specified in the connection to the latest configuration changes
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

            // Retreive JSON Provisioner file and deserialize it
            var filePath = new System.IO.FileInfo(this.ProvisionerFilePath);
            var siteDefinition = JsonConvert.DeserializeObject<SiteProvisionerModel>(System.IO.File.ReadAllText(filePath.FullName));
            var provisionerChoices = siteDefinition.FieldChoices;

            // Load the Context
            var contextWeb = this.ClientContext.Web;
            var fields = this.ClientContext.Web.Fields;
            this.ClientContext.Load(contextWeb, ctxw => ctxw.ServerRelativeUrl, ctxw => ctxw.Id);
            this.ClientContext.Load(fields);

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

            // provision columns
            // Site Columns
            this.ClientContext.Load(fields, f => f.Include(inc => inc.InternalName, inc => inc.JSLink, inc => inc.Title, inc => inc.Id));
            this.ClientContext.ExecuteQueryRetry();

            foreach (var fieldDef in siteDefinition.FieldDefinitions)
            {
                var column = contextWeb.CreateColumn(fieldDef, LogVerbose, LogError, siteGroups, siteDefinition.FieldChoices);
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

            // provision content types
            foreach (var contentDef in siteDefinition.ContentTypes)
            {
                var contentTypeName = contentDef.Name;
                var contentTypeId = contentDef.ContentTypeId;

                if (!contextWeb.ContentTypeExistsById(contentTypeId))
                {
                    contextWeb.CreateContentType(contentTypeName, contentTypeId, (string.IsNullOrEmpty(contentDef.ContentTypeGroup) ? "CustomColumn" : contentDef.ContentTypeGroup));
                }

                foreach (var fieldDef in contentDef.FieldLinkRefs)
                {
                    var siteColumn = siteDefinition.FieldDefinitions.FirstOrDefault(f => f.InternalName == fieldDef);
                    var convertedInternalName = siteColumn.DisplayNameMasked;
                    if (!contextWeb.FieldExistsByNameInContentType(contentTypeName, fieldDef) &&
                        !contextWeb.FieldExistsByNameInContentType(contentTypeName, convertedInternalName))
                    {
                        var column = this.siteColumns.FirstOrDefault(f => f.InternalName == fieldDef);
                        contextWeb.AddFieldToContentTypeByName(contentTypeName, column.FieldGuid, siteColumn.Required);
                    }
                }
            }


            // provision lists
            foreach (var listDef in siteDefinition.Lists
                .Where(w => (string.IsNullOrEmpty(SpecificListName) || (!String.IsNullOrEmpty(SpecificListName) && w.ListName.Equals(SpecificListName, StringComparison.InvariantCultureIgnoreCase)))))
            {
                // Content Type
                var listName = listDef.ListName;
                var listDescription = listDef.ListDescription;


                // provision the list definition
                var siteList = contextWeb.CreateListFromDefinition(listDef, provisionerChoices);


                if (listDef.ContentTypeEnabled && listDef.HasContentTypes)
                {
                    foreach (var contentDef in listDef.ContentTypes)
                    {
                        var contentTypeName = contentDef.Name;
                        ContentType accessContentType = null;

                        if (!contextWeb.ContentTypeExistsByName(listName, contentTypeName))
                        {
                            if (siteDefinition.ContentTypes != null && siteDefinition.ContentTypes.Any(ct => ct.Name == contentTypeName))
                            {
                                contextWeb.AddContentTypeToListByName(listName, contentTypeName, true);

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

                // List Columns
                foreach (var fieldDef in listDef.FieldDefinitions)
                {
                    var column = siteList.CreateListColumn(fieldDef, LogVerbose, LogWarning, siteGroups, provisionerChoices);
                    if (column == null)
                    {
                        LogWarning("Failed to create column {0}.", new string[] { fieldDef.InternalName });
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
                            foreach (var fieldInternalName in contentDef.FieldLinkRefs)
                            {
                                var column = listColumns.FirstOrDefault(f => f.InternalName == fieldInternalName);
                                if (column != null)
                                {
                                    var fieldLinks = accessContentType.FieldLinks;
                                    contextWeb.Context.Load(fieldLinks, cf => cf.Include(inc => inc.Id, inc => inc.Name));
                                    contextWeb.Context.ExecuteQueryRetry();

                                    var convertedInternalName = column.DisplayNameMasked;
                                    if (!fieldLinks.Any(a => a.Name == column.InternalName || a.Name == convertedInternalName))
                                    {
                                        LogVerbose("Content Type {0} Adding Field {1}", new string[] { contentTypeName, column.InternalName });
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
                                else
                                {
                                    LogWarning("Failed to create column {0}.", new string[] { fieldInternalName });
                                }
                            }
                        }
                    }
                }

                // Views
                if (listDef.Views != null && listDef.Views.Any())
                {
                    ViewCollection views = siteList.Views;
                    contextWeb.Context.Load(views, f => f.Include(inc => inc.Id, inc => inc.Hidden, inc => inc.Title, inc => inc.DefaultView));
                    contextWeb.Context.ExecuteQueryRetry();

                    foreach (var viewDef in listDef.Views)
                    {
                        try
                        {
                            if (views.Any(v => v.Title.Equals(viewDef.Title, StringComparison.CurrentCultureIgnoreCase)))
                            {
                                LogVerbose("View {0} found in list {1}", viewDef.Title, listName);
                                continue;
                            }

                            var view = siteList.CreateView(viewDef.InternalName, viewDef.ViewCamlType, viewDef.FieldRefName, viewDef.RowLimit, viewDef.DefaultView, viewDef.QueryXml, viewDef.PersonalView, viewDef.PagedView);
                            contextWeb.Context.Load(view, v => v.Title, v => v.Id, v => v.ServerRelativeUrl);
                            contextWeb.Context.ExecuteQueryRetry();

                            view.Title = viewDef.Title;
                            if (viewDef.HasJsLink && viewDef.JsLink.IndexOf("clienttemplates.js") < -1)
                            {
                                view.JSLink = viewDef.JsLink;
                            }
                            view.Update();
                            contextWeb.Context.ExecuteQueryRetry();
                        }
                        catch (Exception ex)
                        {
                            LogError(ex, "Failed to create view {0} with XML:{1}", viewDef.Title, viewDef.QueryXml);
                        }
                    }
                }


                // List Data if provided in the JSON file, lets add it to the list
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
