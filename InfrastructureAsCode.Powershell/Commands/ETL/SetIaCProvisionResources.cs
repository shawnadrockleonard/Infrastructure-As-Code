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
    [Cmdlet(VerbsCommon.Set, "IaCProvisionResources")]
    [CmdletHelp("Set site definition components based on JSON file.", Category = "ETL")]
    public class SetIaCProvisionResources : IaCCmdlet
    {
        /// <summary>
        /// Represents the directory path for any JSON files for serialization
        /// </summary>
        [Parameter(Mandatory = false)]
        public string SiteContent { get; set; }

        /// <summary>
        /// Validate parameters
        /// </summary>
        protected override void BeginProcessing()
        {
            base.BeginProcessing();
            if (!System.IO.Directory.Exists(this.SiteContent))
            {
                throw new Exception(string.Format("The directory does not exists {0}", this.SiteContent));
            }

            if(!System.IO.Directory.Exists(string.Format("{0}\\Content", this.SiteContent)))
            {
                throw new Exception(string.Format("The content directory does not exists {0}", this.SiteContent));
            }
        }

        internal List<SPGroupDefinitionModel> siteGroups { get; set; }

        internal List<SPFieldDefinitionModel> siteColumns { get; set; }

        /// <summary>
        /// Process the request
        /// </summary>
        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();

            if (this.ClientContext == null)
            {
                LogWarning("Invalid client context, configure the service to run again");
                return;
            }


            siteGroups = new List<SPGroupDefinitionModel>();
            siteColumns = new List<SPFieldDefinitionModel>();

            //Move away from method configuration into a JSON file
            var filePath = string.Format("{0}\\Content\\{1}", this.SiteContent, "Provisioner.json");
            var siteDefinition = JsonConvert.DeserializeObject<SiteProvisionerModel>(System.IO.File.ReadAllText(filePath));

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
            this.ClientContext.ExecuteQuery();

            // creates groups
            siteGroups.AddRange(this.ClientContext.Web.GetOrCreateSiteGroups(siteDefinition.Groups));

            var fileJsonLocation = string.Format("{0}\\Content\\", this.SiteContent);

            foreach (var listDef in siteDefinition.Lists)
            {
                // Content Type
                var listName = listDef.ListName;
                var listDescription = listDef.ListDescription;

                // Site Columns
                FieldCollection fields = this.ClientContext.Web.Fields;
                this.ClientContext.Load(fields, f => f.Include(inc => inc.InternalName, inc => inc.JSLink, inc => inc.Title, inc => inc.Id));
                this.ClientContext.ExecuteQuery();

                // Create List Columns
                foreach (var fieldDef in listDef.FieldDefinitions)
                {
                    var column =  this.ClientContext.Web.CreateColumn(fieldDef, LogVerbose, LogError, siteGroups, fileJsonLocation);
                    if (column == null)
                    {
                        LogWarning("Failed to create column {0}.", fieldDef.InternalName);
                    }
                    else
                    {
                        siteColumns.Add(new SPFieldDefinitionModel()
                        {
                            InternalName = column.InternalName,
                            DisplayName = column.Title,
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
                        this.ClientContext.Web.CreateContentType(contentTypeName, contentTypeId, contentDef.ContentTypeGroup);
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
                    this.ClientContext.ExecuteQuery();

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
                        this.ClientContext.ExecuteQuery();

                        if (listDef.ContentTypeEnabled)
                        {
                            List list = this.ClientContext.Web.GetListByTitle(listName);
                            list.ContentTypesEnabled = true;
                            list.Update();
                            this.ClientContext.Web.Context.ExecuteQuery();
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
                    this.ClientContext.ExecuteQuery();

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
                            this.ClientContext.ExecuteQuery();

                            view.Title = viewDef.Title;
                            view.Update();
                            this.ClientContext.ExecuteQuery();
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
            if (this.ClientContext == null)
            {
                LogWarning("Invalid client context, configure the service to run again");
                return;
            }

            // obtain CSOM object for host web
            this.ClientContext.Load(this.ClientContext.Web, hw => hw.SiteGroups, hw => hw.Title, hw => hw.ContentTypes);
            this.ClientContext.ExecuteQuery();

            // create site groups
            siteGroups.AddRange(this.ClientContext.Web.GetOrCreateSiteGroups(siteDefinition.Groups));

            var fileJsonLocation = string.Format("{0}\\Content\\", this.SiteContent);

            foreach (var listDef in siteDefinition.Lists)
            {
                this.ClientContext.Web.CreateListFromDefinition(listDef, LogVerbose, LogWarning, LogError, siteGroups, fileJsonLocation);
            }
        }
    }
}
