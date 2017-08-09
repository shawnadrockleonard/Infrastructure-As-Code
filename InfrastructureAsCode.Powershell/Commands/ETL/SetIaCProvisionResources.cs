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


        internal List<SPGroupDefinitionModel> siteGroups { get; set; }

        internal List<SPFieldDefinitionModel> siteColumns { get; set; }


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

            if (siteDefinition.SiteResources)
            {
                InstallComponentsToHostWeb(siteDefinition);
            }
            else
            {
                InstallComponentsToHostWebList(siteDefinition);
            }
        }


        /// <summary>
        /// Provision Site Columns
        /// Provision Site Content Type
        /// Provision List
        /// Associate Content Type to List
        /// Provision List Views
        /// </summary>
        public void InstallComponentsToHostWeb(SiteProvisionerModel siteDefinition)
        {

            // obtain CSOM object for host web
            this.ClientContext.Load(this.ClientContext.Web, hw => hw.SiteGroups, hw => hw.Title, hw => hw.ContentTypes);
            this.ClientContext.ExecuteQueryRetry();

            // creates groups
            siteGroups.AddRange(this.ClientContext.Web.GetOrCreateSiteGroups(siteDefinition.Groups));


            foreach (var listDef in siteDefinition.Lists.Where(w =>
                (string.IsNullOrEmpty(SpecificListName) || (!String.IsNullOrEmpty(SpecificListName) && w.ListName.Equals(SpecificListName, StringComparison.InvariantCultureIgnoreCase)))))
            {
                // Content Type
                var listName = listDef.ListName;
                var listDescription = listDef.ListDescription;

                // Site Columns
                FieldCollection fields = this.ClientContext.Web.Fields;
                this.ClientContext.Load(fields, f => f.Include(inc => inc.InternalName, inc => inc.JSLink, inc => inc.Title, inc => inc.Id));
                this.ClientContext.ExecuteQueryRetry();

                // Create List Columns
                foreach (var fieldDef in listDef.FieldDefinitions)
                {
                    var column = this.ClientContext.Web.CreateColumn(fieldDef, LogVerbose, LogError, siteGroups, siteDefinition.FieldChoices);
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

                foreach (var contentDef in listDef.ContentTypes)
                {
                    var contentTypeName = contentDef.Name;
                    var contentTypeId = contentDef.ContentTypeId;

                    if (!this.ClientContext.Web.ContentTypeExistsById(contentTypeId))
                    {
                        this.ClientContext.Web.CreateContentType(contentTypeName, contentTypeId, (string.IsNullOrEmpty(contentDef.ContentTypeGroup) ? "CustomColumn" : contentDef.ContentTypeGroup));
                    }

                    foreach (var fieldDef in contentDef.FieldLinkRefs)
                    {
                        var siteColumn = listDef.FieldDefinitions.FirstOrDefault(f => f.InternalName == fieldDef);
                        var convertedInternalName = siteColumn.DisplayNameMasked;
                        if (!this.ClientContext.Web.FieldExistsByNameInContentType(contentTypeName, fieldDef) &&
                            !this.ClientContext.Web.FieldExistsByNameInContentType(contentTypeName, convertedInternalName))
                        {
                            var column = this.siteColumns.FirstOrDefault(f => f.InternalName == fieldDef);
                            this.ClientContext.Web.AddFieldToContentTypeByName(contentTypeName, column.FieldGuid, siteColumn.Required);
                        }
                    }

                    // check to see if Picture library named Photos already exists
                    ListCollection allLists = this.ClientContext.Web.Lists;
                    IEnumerable<List> foundLists = this.ClientContext.LoadQuery(allLists.Where(list => list.Title == listName));
                    this.ClientContext.ExecuteQueryRetry();

                    List accessRequest = foundLists.FirstOrDefault();
                    if (accessRequest == null)
                    {
                        // create Picture library named Photos if it does not already exist
                        ListCreationInformation accessRequestInfo = new ListCreationInformation();
                        accessRequestInfo.Title = listName;
                        accessRequestInfo.Description = listDescription;
                        accessRequestInfo.QuickLaunchOption = listDef.QuickLaunch;
                        accessRequestInfo.TemplateType = (int)listDef.ListTemplate;
                        accessRequestInfo.Url = listName;
                        accessRequest = this.ClientContext.Web.Lists.Add(accessRequestInfo);
                        this.ClientContext.ExecuteQueryRetry();

                        if (listDef.ContentTypeEnabled)
                        {
                            List list = this.ClientContext.Web.GetListByTitle(listName);
                            list.ContentTypesEnabled = true;
                            list.Update();
                            this.ClientContext.Web.Context.ExecuteQueryRetry();
                        }
                    }

                    if (listDef.ContentTypeEnabled)
                    {
                        if (!this.ClientContext.Web.ContentTypeExistsByName(listName, contentTypeName))
                        {
                            this.ClientContext.Web.AddContentTypeToListByName(listName, contentTypeName, true);
                        }

                        // Set the content type as default content type to the TestLib list
                        //this.ClientContext.Web.SetDefaultContentTypeToList(listName, contentTypeId);
                    }

                    // Views
                    ViewCollection views = accessRequest.Views;
                    this.ClientContext.Load(views, f => f.Include(inc => inc.Id, inc => inc.Hidden, inc => inc.Title, inc => inc.DefaultView));
                    this.ClientContext.ExecuteQueryRetry();

                    foreach (var viewDef in listDef.Views)
                    {
                        try
                        {
                            if (views.Any(v => v.Title == viewDef.Title))
                            {
                                LogVerbose("View {0} found in list {1}", viewDef.Title, listName);
                                continue;
                            }

                            var view = accessRequest.CreateView(viewDef.InternalName, viewDef.ViewCamlType, viewDef.FieldRefName, viewDef.RowLimit, viewDef.DefaultView, viewDef.QueryXml);
                            this.ClientContext.Load(view, v => v.Title, v => v.Id, v => v.ServerRelativeUrl);
                            this.ClientContext.ExecuteQueryRetry();

                            view.Title = viewDef.Title;
                            view.Update();
                            this.ClientContext.ExecuteQueryRetry();
                        }
                        catch (Exception ex)
                        {
                            LogError(ex, "Failed to create view {0} with XML:{1}", viewDef.Title, viewDef.QueryXml);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Provision List
        /// Provision Content Type to List
        /// Provision Field to Content Type
        /// Provision List Views
        /// </summary>
        public void InstallComponentsToHostWebList(SiteProvisionerModel siteDefinition)
        {
            // obtain CSOM object for host web
            this.ClientContext.Load(this.ClientContext.Web, hw => hw.SiteGroups, hw => hw.Title, hw => hw.ContentTypes);
            this.ClientContext.ExecuteQueryRetry();

            // create site groups
            siteGroups.AddRange(this.ClientContext.Web.GetOrCreateSiteGroups(siteDefinition.Groups));

            foreach (var listDef in siteDefinition.Lists.Where(w =>
                (string.IsNullOrEmpty(SpecificListName) || (!String.IsNullOrEmpty(SpecificListName) && w.ListName.Equals(SpecificListName, StringComparison.InvariantCultureIgnoreCase)))))
            {
                this.ClientContext.Web.CreateListFromDefinition(listDef, LogVerbose, LogWarning, LogError, siteGroups, siteDefinition.FieldChoices);
            }
        }
    }
}
