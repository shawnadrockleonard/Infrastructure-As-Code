using InfrastructureAsCode.Core.Extensions;
using InfrastructureAsCode.Core.Models;
using InfrastructureAsCode.Core.Models.Minimal;
using InfrastructureAsCode.Powershell.CmdLets;
using InfrastructureAsCode.Powershell.PipeBinds;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Utilities;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace InfrastructureAsCode.Powershell.Commands.Lists
{
    /// <summary>
    /// Returns the list definition, views, columns, settings
    /// </summary>
    /// <remarks>
    /// Get-IaCListDefinition -List ""Demo List""
    /// </remarks>
    [Cmdlet(VerbsCommon.Get, "IaCListDefinition")]
    [OutputType(typeof(SPListDefinition))]
    public class GetIaCListDefinition : IaCCmdlet
    {
        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0, HelpMessage = "The ID or Url of the list.")]
        public ListPipeBind Identity;

        /// <summary>
        /// Expand the list definition
        /// </summary>
        [Parameter(Mandatory = false, Position = 1)]
        public SwitchParameter ExpandObjects { get; set; }

        /// <summary>
        /// Extract the list items
        /// </summary>
        [Parameter(Mandatory = false, Position = 2)]
        public SwitchParameter ExtractData { get; set; }


        public override void ExecuteCmdlet()
        {

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

            // Construct the model
            var SiteComponents = new SiteProvisionerModel();


            if (Identity != null)
            {
                var list = Identity.GetList(this.ClientContext.Web);
                if (list != null)
                {
                    // We'll focus on the List Definition and not Site elements
                    SiteComponents.FieldChoices = new List<SiteProvisionerFieldChoiceModel>();
                    SiteComponents.Lists = new List<SPListDefinition>();

                    // ---> Site Usage Properties
                    var _ctx = this.ClientContext;
                    var _contextWeb = this.ClientContext.Web;
                    var _site = this.ClientContext.Site;

                    ClientContext.Load(_contextWeb, ctxw => ctxw.ServerRelativeUrl, ctxw => ctxw.Id);
                    ClientContext.Load(_site, cts => cts.Usage, cts => cts.Id);

                    ClientContext.Load(list,
                        lctx => lctx.Id,
                        lctx => lctx.Title,
                        lctx => lctx.Description,
                        lctx => lctx.DefaultViewUrl,
                        lctx => lctx.Hidden,
                        lctx => lctx.IsApplicationList,
                        lctx => lctx.IsCatalog,
                        lctx => lctx.IsSiteAssetsLibrary,
                        lctx => lctx.IsPrivate,
                        lctx => lctx.IsSystemList,
                        lctx => lctx.Created,
                        lctx => lctx.LastItemModifiedDate,
                        lctx => lctx.LastItemUserModifiedDate,
                        lctx => lctx.OnQuickLaunch,
                        lctx => lctx.ContentTypesEnabled,
                        lctx => lctx.EnableFolderCreation,
                        lctx => lctx.EnableModeration,
                        lctx => lctx.EnableVersioning,
                        lctx => lctx.CreatablesInfo,
                        lctx => lctx.EnableVersioning,
                        lctx => lctx.RootFolder.ServerRelativeUrl);
                    ClientContext.ExecuteQueryRetry();



                    var weburl = TokenHelper.EnsureTrailingSlash(_contextWeb.ServerRelativeUrl);


                    var listmodel = new SPListDefinition()
                    {
                        Id = list.Id,
                        ListName = list.Title,
                        ServerRelativeUrl = list.DefaultViewUrl,
                        Created = list.Created,
                        LastItemModifiedDate = list.LastItemModifiedDate,
                        LastItemUserModifiedDate = list.LastItemUserModifiedDate,
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


                    if (ExpandObjects)
                    {
                        var listurl = TokenHelper.EnsureTrailingSlash(list.RootFolder.ServerRelativeUrl);

                        var views = ClientContext.LoadQuery(list.Views
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
                                ));


                        var fields = ClientContext.LoadQuery(list.Fields
                            .Include(
                                v => v.AutoIndexed,
                                v => v.CanBeDeleted,
                                v => v.DefaultFormula,
                                v => v.DefaultValue,
                                v => v.Description,
                                v => v.EnforceUniqueValues,
                                v => v.FieldTypeKind,
                                v => v.Filterable,
                                v => v.Group,
                                v => v.Hidden,
                                v => v.Id,
                                v => v.InternalName,
                                v => v.Indexed,
                                v => v.JSLink,
                                v => v.NoCrawl,
                                v => v.ReadOnlyField,
                                v => v.Required,
                                v => v.Title,
                                v => v.SchemaXml));
                        ClientContext.ExecuteQueryRetry();


                        listmodel.Views = new List<SPViewDefinitionModel>();
                        listmodel.InternalViews = new List<SPViewDefinitionModel>();

                        foreach (var view in views)
                        {
                            ViewType viewCamlType = ViewType.None;
                            foreach (var vtype in Enum.GetNames(typeof(ViewType)))
                            {
                                if (vtype.Equals(view.ViewType, StringComparison.InvariantCultureIgnoreCase))
                                {
                                    viewCamlType = (ViewType)Enum.Parse(typeof(ViewType), vtype);
                                    break;
                                }
                            }

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

                            var vinternal = (view.ServerRelativeUrl.IndexOf(listurl, StringComparison.CurrentCultureIgnoreCase) == -1);
                            if (vinternal)
                            {
                                viewmodel.SitePage = view.ServerRelativeUrl.Replace(weburl, "");
                                viewmodel.InternalView = true;
                            }
                            else
                            {
                                viewmodel.InternalName = view.ServerRelativeUrl.Replace(listurl, "").Replace(".aspx", "");
                            }

                            foreach (var vfields in view.ViewFields)
                            {
                                viewmodel.FieldRefName.Add(vfields);
                            }

                            var vjslinks = view.JSLink.Split(new string[] { "|" }, StringSplitOptions.RemoveEmptyEntries);
                            if (vjslinks != null && !vjslinks.Any(jl => jl == "clienttemplates.js"))
                            {
                                viewmodel.JsLinkFiles = new List<string>();
                                foreach (var vjslink in vjslinks)
                                {
                                    viewmodel.JsLinkFiles.Add(vjslink);
                                }
                            }

                            if (view.Hidden)
                            {
                                listmodel.InternalViews.Add(viewmodel);
                            }
                            else
                            {
                                listmodel.Views.Add(viewmodel);
                            }
                        }


                        foreach (var listField in fields)
                        {
                            // skip internal fields
                            if (skiptypes.Any(st => listField.FieldTypeKind == st))
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
                                        listmodel.FieldDefinitions.Add(listField.RetrieveField(_contextWeb, null, xField));
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Trace.TraceError("Failed to parse field {0} MSG:{1}", listField.InternalName, ex.Message);
                            }
                        }
                    }

                    if (ExtractData)
                    {
                        listmodel.ListItems = new List<SPListItemDefinition>();

                        ListItemCollectionPosition itemPosition = null;
                        var camlQuery = new CamlQuery()
                        {
                            ViewXml = CAML.ViewQuery(ViewScope.RecursiveAll,
                                        string.Empty,
                                        CAML.OrderBy(new OrderByField("Title")),
                                        CAML.ViewFields((new string[] { "Title", "Author", "Created", "Editor", "Modified" }).Select(s => CAML.FieldRef(s)).ToArray()),
                                        50)
                        };

                        try
                        {
                            while (true)
                            {
                                camlQuery.ListItemCollectionPosition = itemPosition;
                                ListItemCollection listItems = list.GetItems(camlQuery);
                                this.ClientContext.Load(listItems);
                                this.ClientContext.ExecuteQueryRetry();
                                itemPosition = listItems.ListItemCollectionPosition;

                                foreach (var rbiItem in listItems)
                                {

                                    LogVerbose("Title: {0}; Item ID: {1}", rbiItem["Title"], rbiItem.Id);

                                    var author = rbiItem.RetrieveListItemUserValue("Author");
                                    var editor = rbiItem.RetrieveListItemUserValue("Editor");

                                    var newitem = new SPListItemDefinition()
                                    {
                                        Title = rbiItem.RetrieveListItemValue("Title"),
                                        Id = rbiItem.Id,
                                        Created = rbiItem.RetrieveListItemValue("Created").ToNullableDatetime(),
                                        CreatedBy = new SPPrincipalUserDefinition()
                                        {
                                            Id = author.LookupId,
                                            LoginName = author.LookupValue,
                                            Email = author.Email
                                        },
                                        Modified = rbiItem.RetrieveListItemValue("Modified").ToNullableDatetime(),
                                        ModifiedBy = new SPPrincipalUserDefinition()
                                        {
                                            Id = editor.LookupId,
                                            LoginName = editor.LookupValue,
                                            Email = editor.Email
                                        }
                                    };

                                    listmodel.ListItems.Add(newitem);
                                }

                                if (itemPosition == null)
                                {
                                    break;
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            LogError(ex, "Failed to retrieve list item collection");
                        }
                    }

                   SiteComponents.Lists.Add(listmodel);
                }
            }

            // Write the model to memory
            WriteObject(SiteComponents);
        }
    }
}
