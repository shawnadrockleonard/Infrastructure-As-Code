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
                    lctx => lctx.SchemaXml,
                    linc => linc.Views
                    .Include(
                        lvt => lvt.Title,
                        lvt => lvt.DefaultView,
                        lvt => lvt.ServerRelativeUrl,
                        lvt => lvt.Id,
                        lvt => lvt.Aggregations,
                        lvt => lvt.AggregationsStatus,
                        lvt => lvt.BaseViewId,
                        lvt => lvt.Hidden,
                        lvt => lvt.ImageUrl,
                        lvt => lvt.JSLink,
                        lvt => lvt.HtmlSchemaXml,
                        lvt => lvt.ListViewXml,
                        lvt => lvt.MobileDefaultView,
                        lvt => lvt.ModerationType,
                        lvt => lvt.OrderedView,
                        lvt => lvt.Paged,
                        lvt => lvt.PageRenderType,
                        lvt => lvt.PersonalView,
                        lvt => lvt.ReadOnlyView,
                        lvt => lvt.Scope,
                        lvt => lvt.RowLimit,
                        lvt => lvt.StyleId,
                        lvt => lvt.TabularView,
                        lvt => lvt.Threaded,
                        lvt => lvt.Toolbar,
                        lvt => lvt.ToolbarTemplateName,
                        lvt => lvt.ViewFields,
                        lvt => lvt.ViewJoins,
                        lvt => lvt.ViewQuery,
                        lvt => lvt.ViewType,
                        lvt => lvt.ViewProjectedFields,
                        lvt => lvt.Method
                        ),
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
                    LogVerbose("Processing list {0}", list.Title);

                    var listdefinition = new SPListDefinition()
                    {
                        Id = list.Id,
                        ListName = list.Title,
                        ListDescription = list.Description,
                        QuickLaunch = list.OnQuickLaunch ? QuickLaunchOptions.On : QuickLaunchOptions.Off,
                        ContentTypeEnabledOverride = list.ContentTypesEnabled,
                        EnableFolderCreation = list.EnableFolderCreation,
                        Hidden = list.Hidden,
                        IsApplicationList = list.IsApplicationList,
                        IsCatalog = list.IsCatalog,
                        IsSiteAssetsLibrary = list.IsSiteAssetsLibrary,
                        IsPrivate = list.IsPrivate,
                        IsSystemList = list.IsSystemList
                    };


                    if (list.ContentTypes != null && list.ContentTypes.Any())
                    {
                        listdefinition.ContentTypes = new List<SPContentTypeDefinition>();
                        foreach (var contenttype in list.ContentTypes)
                        {
                            LogVerbose("Processing list {0} content type {1}", list.Title, contenttype.Name);

                            var ctypemodel = new SPContentTypeDefinition()
                            {
                                Inherits = true,
                                ContentTypeId = contenttype.Id.StringValue,
                                ContentTypeGroup = contenttype.Group,
                                Description = contenttype.Description,
                                Name = contenttype.Name,
                                Hidden = contenttype.Hidden,
                                JSLink = contenttype.JSLink
                            };

                            if (contenttype.FieldLinks.Any())
                            {
                                ctypemodel.FieldLinks = new List<SPFieldLinkDefinitionModel>();
                                foreach (var cfield in contenttype.FieldLinks)
                                {
                                    ctypemodel.FieldLinks.Add(new SPFieldLinkDefinitionModel()
                                    {
                                        Id = cfield.Id,
                                        Name = cfield.Name,
                                        Hidden = cfield.Hidden,
                                        Required = cfield.Required
                                    });

                                }
                            }

                            if (contenttype.Fields.Any())
                            {
                                foreach (var cfield in contenttype.Fields.Where(cf => !ctypemodel.FieldLinks.Any(fl => fl.Name == cf.InternalName)))
                                {
                                    ctypemodel.FieldLinks.Add(new SPFieldLinkDefinitionModel()
                                    {
                                        Id = cfield.Id,
                                        Name = cfield.InternalName,
                                        Hidden = cfield.Hidden,
                                        Required = cfield.Required
                                    });
                                }
                            }

                            listdefinition.ContentTypes.Add(ctypemodel);
                        }
                    }

                    var listfields = new List<SPFieldDefinitionModel>();
                    if (list.Fields != null && list.Fields.Any())
                    {
                        foreach (Field listField in list.Fields)
                        {
                            LogVerbose("Processing list {0} field {1}", list.Title, listField.InternalName);

                            // skip internal fields
                            if (skiptypes.Any(st => listField.FieldTypeKind == st)
                                || skipcolumns.Any(sg => listField.Group.Equals(sg, StringComparison.CurrentCultureIgnoreCase)))
                            {
                                continue;
                            }

                            // skip fields that are defined
                            if (contentTypesFieldset.Any(ft => ft.name == listField.InternalName))
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
                                    if (xSourceID != null && xSourceID.Value.IndexOf(ns.NamespaceName, StringComparison.CurrentCultureIgnoreCase) < 0)
                                    {
                                        var customField = listField.RetrieveField(contextWeb, groupQuery, xField);
                                        listfields.Add(customField);

                                        if (customField.FieldTypeKind == FieldType.Lookup)
                                        {
                                            listdefinition.ListDependency.Add(customField.LookupListName);
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Trace.TraceError("Failed to parse field {0} MSG:{1}", listField.InternalName, ex.Message);
                            }
                        }

                        listdefinition.FieldDefinitions = listfields;
                    }

                    if (list.Views != null && list.Views.Any())
                    {
                        listdefinition.InternalViews = new List<SPViewDefinitionModel>();
                        listdefinition.Views = new List<SPViewDefinitionModel>();
                        var listurl = TokenHelper.EnsureTrailingSlash(list.RootFolder.ServerRelativeUrl);

                        foreach (var view in list.Views)
                        {
                            LogVerbose("Processing list {0} view {1}", list.Title, view.Title);

                            var vinternal = (view.ServerRelativeUrl.IndexOf(listurl, StringComparison.CurrentCultureIgnoreCase) == -1);

                            ViewType viewCamlType = InfrastructureAsCode.Core.Extensions.ListExtensions.TryGetViewType(view.ViewType);

                            var viewmodel = new SPViewDefinitionModel()
                            {
                                Id = view.Id,
                                Title = view.Title,
                                DefaultView = view.DefaultView,
                                FieldRefName = new List<string>(),
                                Aggregations = view.Aggregations,
                                AggregationsStatus = view.AggregationsStatus,
                                BaseViewId = view.BaseViewId,
                                Hidden = view.Hidden,
                                ImageUrl = view.ImageUrl,
                                Toolbar = view.Toolbar,
                                ListViewXml = view.ListViewXml,
                                MobileDefaultView = view.MobileDefaultView,
                                ModerationType = view.ModerationType,
                                OrderedView = view.OrderedView,
                                Paged = view.Paged,
                                PageRenderType = view.PageRenderType,
                                PersonalView = view.PersonalView,
                                ReadOnlyView = view.ReadOnlyView,
                                Scope = view.Scope,
                                RowLimit = view.RowLimit,
                                StyleId = view.StyleId,
                                TabularView = view.TabularView,
                                Threaded = view.Threaded,
                                ViewJoins = view.ViewJoins,
                                ViewQuery = view.ViewQuery,
                                ViewCamlType = viewCamlType
                            };

                            if (vinternal)
                            {
                                viewmodel.SitePage = view.ServerRelativeUrl.Replace(weburl, "");
                                viewmodel.InternalView = true;
                            }
                            else
                            {
                                viewmodel.InternalName = view.ServerRelativeUrl.Replace(listurl, "").Replace(".aspx", "");
                            }

                            if (view.ViewFields != null && view.ViewFields.Any())
                            {
                                foreach (var vfields in view.ViewFields)
                                {
                                    viewmodel.FieldRefName.Add(vfields);
                                }
                            }

                            if (view.JSLink != null && view.JSLink.Any())
                            {
                                var vjslinks = view.JSLink.Split(new string[] { "|" }, StringSplitOptions.RemoveEmptyEntries);
                                if (vjslinks != null && !vjslinks.Any(jl => jl == "clienttemplates.js"))
                                {
                                    viewmodel.JsLinkFiles = new List<string>();
                                    foreach (var vjslink in vjslinks)
                                    {
                                        viewmodel.JsLinkFiles.Add(vjslink);
                                    }
                                }
                            }

                            if (view.Hidden)
                            {
                                listdefinition.InternalViews.Add(viewmodel);
                            }
                            else
                            {
                                listdefinition.Views.Add(viewmodel);
                            }
                        }
                    }

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

                        if (listPopped.ListDependency.Any(listField =>
                                !SiteComponents.Lists.Any(sc => sc.ListName.Equals(listField, StringComparison.InvariantCultureIgnoreCase)
                                                             || sc.InternalName.Equals(listField, StringComparison.InvariantCultureIgnoreCase))))
                        {
                            // no list definition exists in the collection with the dependent lookup lists
                            LogWarning("List {0} depends on {1} which do not exist current collection", listPopped.ListName, string.Join(",", listPopped.ListDependency));
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
