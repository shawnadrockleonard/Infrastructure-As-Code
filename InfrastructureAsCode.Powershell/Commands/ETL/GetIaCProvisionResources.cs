using InfrastructureAsCode.Core.Models;
using InfrastructureAsCode.Powershell.CmdLets;
using InfrastructureAsCode.Core.Extensions;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using OfficeDevPnP.Core.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using InfrastructureAsCode.Core.Reports;

namespace InfrastructureAsCode.Powershell.Commands.ETL
{
    /// <summary>
    /// The function cmdlet will query the site specified in the connection and build a configuration file
    /// </summary>
    [Cmdlet(VerbsCommon.Get, "IaCProvisionResources")]
    [CmdletHelp("Get site definition components and write to JSON file.", Category = "ETL")]
    public class GetIaCProvisionResources : IaCCmdlet
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
        public string[] SpecificLists { get; set; }

        private bool _filterLists { get; set; }

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
            var fileinfo = new System.IO.FileInfo(ProvisionerFilePath);

            if (!fileinfo.Directory.Exists)
            {
                throw new System.IO.DirectoryNotFoundException(string.Format("The provisioner directory was not found {0}", fileinfo.DirectoryName));
            }

            _filterLists = (SpecificLists != null && SpecificLists.Any());
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

            // File Info
            var fileInfo = new System.IO.FileInfo(this.ProvisionerFilePath);

            // Initialize logging instance with Powershell logger
            ITraceLogger logger = new DefaultUsageLogger(LogVerbose, LogWarning, LogError);

            // SharePoint URI for XML parsing
            XNamespace ns = "http://schemas.microsoft.com/sharepoint/";

            // Skip these specific fields
            var skiptypes = new FieldType[]
            {
                FieldType.Computed,
                FieldType.ContentTypeId,
                FieldType.Invalid,
                FieldType.WorkflowStatus,
                FieldType.WorkflowEventType,
                FieldType.Threading,
                FieldType.ThreadIndex,
                FieldType.Recurrence,
                FieldType.PageSeparator,
                FieldType.OutcomeChoice
            };

            var skipcolumns = new string[]
            {
                "_Hidden",
                "Base Columns",
                "Content Feedback",
                "Core Contact and Calendar Columns",
                "Core Document Columns",
                "Core Task and Issue Columns",
                "Display Template Columns",
                "Document and Record Management Columns",
                "Enterprise Keywords Group",
                "Extended Columns",
                "JavaScript Display Template Columns",
                "Page Layout Columns",
                "Publishing Columns",
                "Reports",
                "Status Indicators",
                "Translation Columns"
            };

            var skipcontenttypes = new string[]
            {
                "_Hidden",
                "Business Intelligence",
                "Community Content Types",
                "Digital Asset Content Types",
                "Display Template Content Types",
                "Document Content Types",
                "Document Set Content Types",
                "Folder Content Types",
                "Content Feedback",
                "Publishing Content Types",
                "Page Layout Content Types",
                "Special Content Types",
                "Group Work Content Types",
                "List Content Types"
            };

            // Construct the model
            var SiteComponents = new SiteProvisionerModel();

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
                    linc => linc.RootFolder.ServerRelativeUrl,
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
                    lctx => lctx.SchemaXml).Where(w => !w.IsSystemList && !w.IsSiteAssetsLibrary));
            this.ClientContext.ExecuteQueryRetry();

            var weburl = TokenHelper.EnsureTrailingSlash(contextWeb.ServerRelativeUrl);


            if (groupQuery.Any())
            {
                SiteComponents.Groups = new List<SPGroupDefinitionModel>();

                foreach (var group in groupQuery)
                {
                    var model = new SPGroupDefinitionModel()
                    {
                        Id = group.Id,
                        Title = group.Title,
                        Description = group.Description,
                        AllowRequestToJoinLeave = group.AllowRequestToJoinLeave,
                        AllowMembersEditMembership = group.AllowMembersEditMembership,
                        AutoAcceptRequestToJoinLeave = group.AutoAcceptRequestToJoinLeave,
                        OnlyAllowMembersViewMembership = group.OnlyAllowMembersViewMembership,
                        RequestToJoinLeaveEmailSetting = group.RequestToJoinLeaveEmailSetting
                    };

                    SiteComponents.Groups.Add(model);
                }
            }



            if (fields.Any())
            {
                var webfields = new List<SPFieldDefinitionModel>();
                foreach (Microsoft.SharePoint.Client.Field field in fields)
                {
                    if (skiptypes.Any(st => field.FieldTypeKind == st)
                        || skipcolumns.Any(sg => field.Group.Equals(sg, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        continue;
                    }

                    try
                    {
                        var fieldModel = field.RetrieveField(contextWeb, groupQuery);
                        webfields.Add(fieldModel);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Trace.TraceError("Failed to parse field {0} MSG:{1}", field.InternalName, ex.Message);
                    }
                }

                SiteComponents.FieldDefinitions = webfields;
            }


            var contentTypesFieldset = new List<dynamic>();
            if (contentTypes.Any())
            {
                SiteComponents.ContentTypes = new List<SPContentTypeDefinition>();
                foreach (ContentType contenttype in contentTypes)
                {
                    // skip core content types
                    if (skipcontenttypes.Any(sg => contenttype.Group.Equals(sg, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        continue;
                    }

                    contenttype.EnsureProperties(ctp => ctp.Id, ctp => ctp.Group, ctp => ctp.Hidden, ctp => ctp.Description, ctp => ctp.Name, ctp => ctp.FieldLinks);

                    var ctypemodel = new SPContentTypeDefinition()
                    {
                        ContentTypeId = contenttype.Id.StringValue,
                        ContentTypeGroup = contenttype.Group,
                        Hidden = contenttype.Hidden,
                        Description = contenttype.Description,
                        Name = contenttype.Name
                    };

                    if (contenttype.FieldLinks.Any())
                    {
                        ctypemodel.FieldLinks = new List<SPFieldLinkDefinitionModel>();
                        foreach (FieldLink fieldlink in contenttype.FieldLinks)
                        {
                            ctypemodel.FieldLinks.Add(new SPFieldLinkDefinitionModel()
                            {
                                Id = fieldlink.Id,
                                Name = fieldlink.Name,
                                Required = fieldlink.Required,
                                Hidden = fieldlink.Hidden
                            });

                            contentTypesFieldset.Add(new { ctypeid = contenttype.Id.StringValue, name = fieldlink.Name });
                        }
                    }

                    SiteComponents.ContentTypes.Add(ctypemodel);
                }
            }


            if (lists.Any())
            {
                var sitelists = new List<SPListDefinition>();

                foreach (List list in lists.Where(lwt =>
                    (!_filterLists
                        || (_filterLists && SpecificLists.Any(sl => lwt.Title.Equals(sl, StringComparison.InvariantCultureIgnoreCase))))))
                {
                    var listdefinition = contextWeb.GetListDefinition(list, true, logger, skiptypes, groupQuery);
                    sitelists.Add(listdefinition);
                }

                if (sitelists.Any())
                {
                    var idx = 0;
                    SiteComponents.Lists = new List<SPListDefinition>();

                    // lets add any list with NO lookups first
                    var nolookups = sitelists.Where(sl => !sl.ListDependency.Any());
                    nolookups.ToList().ForEach(nolookup =>
                    {
                        LogVerbose("adding list {0}", nolookup.ListName);
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
                        LogVerbose("adding list {0}", listPopped.ListName);

                        if(listPopped.ListDependency.Any(listField => listPopped.ListName.Equals(listField, StringComparison.InvariantCultureIgnoreCase)
                                                             || listPopped.InternalName.Equals(listField, StringComparison.InvariantCultureIgnoreCase)))
                        {
                            // the list has a dependency on itself
                            LogVerbose("Adding list {0} with dependency on itself", listPopped.ListName);
                            listPopped.ProvisionOrder = idx++;
                            SiteComponents.Lists.Add(listPopped);
                        }
                        else if (listPopped.ListDependency.Any(listField =>
                                !SiteComponents.Lists.Any(sc => sc.ListName.Equals(listField, StringComparison.InvariantCultureIgnoreCase)
                                                             || sc.InternalName.Equals(listField, StringComparison.InvariantCultureIgnoreCase))))
                        {
                            // no list definition exists in the collection with the dependent lookup lists
                            LogWarning("List {0} depends on {1} which do not exist current collection", listPopped.ListName, string.Join(",", listPopped.ListDependency));
                            foreach (var dependent in listPopped.ListDependency.Where(dlist => !haslookups.Any(hl => hl.ListName == dlist)))
                            {
                                var sitelist = contextWeb.GetListByTitle(dependent, lctx => lctx.Id, lctx => lctx.Title, lctx => lctx.RootFolder.ServerRelativeUrl);
                                var listDefinition = contextWeb.GetListDefinition(sitelist, true, logger, skiptypes, groupQuery);
                                haslookups.Add(listDefinition);
                            }
                            
                            haslookups.Add(listPopped); // add back to collection
                            listPopped = null;
                        }
                        else
                        {
                            LogVerbose("Adding list {0} to collection", listPopped.ListName);
                            listPopped.ProvisionOrder = idx++;
                            SiteComponents.Lists.Add(listPopped);
                        }
                    }


                }
            }

            // Write the JSON to disc
            var jsonsettings = new JsonSerializerSettings()
            {
                Formatting = Formatting.Indented,
                Culture = System.Globalization.CultureInfo.CurrentUICulture,
                DateFormatHandling = DateFormatHandling.IsoDateFormat,
                NullValueHandling = NullValueHandling.Ignore
            };

            var json = JsonConvert.SerializeObject(SiteComponents, jsonsettings);
            System.IO.File.WriteAllText(fileInfo.FullName, json);
        }

    }
}
