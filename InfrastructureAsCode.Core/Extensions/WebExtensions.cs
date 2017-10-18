using InfrastructureAsCode.Core.Constants;
using InfrastructureAsCode.Core.Models;
using InfrastructureAsCode.Core.Reports;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WebParts;
using OfficeDevPnP.Core.Entities;
using OfficeDevPnP.Core.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;

namespace InfrastructureAsCode.Core.Extensions
{
    public static class WebExtensions
    {
        public static Web GetAssociatedWeb(this SecurableObject securable)
        {
            if (securable is Web)
            {
                return (Web)securable;
            }

            if (securable is List)
            {
                var list = (List)securable;
                var web = list.ParentWeb;
                securable.Context.Load(web);
                securable.Context.ExecuteQueryRetry();

                return web;
            }

            if (securable is ListItem)
            {
                var listItem = (ListItem)securable;
                var web = listItem.ParentList.ParentWeb;
                securable.Context.Load(web);
                securable.Context.ExecuteQueryRetry();

                return web;
            }

            throw new Exception("Only Web, List, ListItem supported as SecurableObjects");
        }

        public static Web GetWebById(this Web currentWeb, Guid guid)
        {
            var clientContext = currentWeb.Context as ClientContext;
            Site site = clientContext.Site;
            Web web = site.OpenWebById(guid);
            clientContext.Load(web, w => w.Url, w => w.Title, w => w.Id);
            clientContext.ExecuteQueryRetry();

            return web;
        }

        public static Web GetWebByUrl(this Web currentWeb, string url)
        {
            var clientContext = currentWeb.Context as ClientContext;
            Site site = clientContext.Site;
            Web web = site.OpenWeb(url);
            clientContext.Load(web, w => w.Url, w => w.Title, w => w.Id);
            clientContext.ExecuteQueryRetry();

            return web;
        }

        /// <summary>
        /// Will scan the web site groups for groups to provision and retrieve
        /// </summary>
        /// <param name="hostWeb">The host web to which the groups will be created</param>
        /// <param name="groupDef">The collection of groups to retrieve and/or provision</param>
        /// <returns></returns>
        public static Microsoft.SharePoint.Client.Group GetOrCreateSiteGroups(this Web hostWeb, SPGroupDefinitionModel groupDef)
        {
            hostWeb.EnsureProperties(hw => hw.CurrentUser);
            var context = hostWeb.Context;

            // Create Group
            var groupCreationInfo = new GroupCreationInformation();
            groupCreationInfo.Title = groupDef.Title;
            groupCreationInfo.Description = groupDef.Description;

            var oGroup = hostWeb.SiteGroups.Add(groupCreationInfo);
            context.Load(oGroup);

            oGroup.Owner = hostWeb.CurrentUser;
            oGroup.OnlyAllowMembersViewMembership = groupDef.OnlyAllowMembersViewMembership;
            oGroup.AllowMembersEditMembership = groupDef.AllowMembersEditMembership;
            oGroup.AllowRequestToJoinLeave = groupDef.AllowRequestToJoinLeave;
            oGroup.AutoAcceptRequestToJoinLeave = groupDef.AutoAcceptRequestToJoinLeave;
            oGroup.Update();

            context.Load(oGroup, g => g.Id, g => g.Title);
            context.ExecuteQueryRetry();


            return oGroup;
        }

        /// <summary>
        /// Provisions a site column based on the field definition specified
        /// </summary>
        /// <param name="hostWeb">The instantiated site/web to which the column will be retrieved and/or provisioned</param>
        /// <param name="fieldDefinition">The definition for the field</param>
        /// <param name="loggerVerbose">Provides a method for verbose logging</param>
        /// <param name="loggerError">Provides a method for exception logging</param>
        /// <param name="SiteGroups">(OPTIONAL) collection of group, required if this is a PeoplePicker field</param>
        /// <param name="provisionerChoices">(OPTIONAL) deserialized choices from JSON</param>
        /// <returns></returns>
        public static Field CreateColumn(this Web hostWeb, SPFieldDefinitionModel fieldDefinition, ITraceLogger logger, List<SPGroupDefinitionModel> SiteGroups, List<SiteProvisionerFieldChoiceModel> provisionerChoices = null)
        {
            if (fieldDefinition == null)
            {
                throw new ArgumentNullException("fieldDef", "Field definition is required.");
            }

            if (string.IsNullOrEmpty(fieldDefinition.InternalName))
            {
                throw new ArgumentNullException("InternalName");
            }

            if (fieldDefinition.LoadFromJSON && (provisionerChoices == null || !provisionerChoices.Any(pc => pc.FieldInternalName == fieldDefinition.InternalName)))
            {
                throw new ArgumentNullException("provisionerChoices", string.Format("You must specify a collection of field choices for the field {0}", fieldDefinition.Title));
            }

            var fields = hostWeb.Context.LoadQuery(hostWeb.Fields.Include(f => f.Id, f => f.InternalName, f => f.Title, f => f.JSLink, f => f.Indexed, f => f.CanBeDeleted, f => f.Required));
            hostWeb.Context.ExecuteQueryRetry();


            var returnField = fields.FirstOrDefault(f => f.Id == fieldDefinition.FieldGuid || f.InternalName == fieldDefinition.InternalName || f.Title == fieldDefinition.InternalName);
            if (returnField == null)
            {
                var finfoXml = hostWeb.CreateFieldDefinition(fieldDefinition, SiteGroups, provisionerChoices);
                logger.LogInformation("Provision Site field {0} with XML:{1}", fieldDefinition.InternalName, finfoXml);
                try
                {
                    var createdField = hostWeb.CreateField(finfoXml, executeQuery: false);
                    hostWeb.Context.ExecuteQueryRetry();
                }
                catch (Exception ex)
                {
                    logger.LogError(ex, "EXCEPTION: field {0} with message {1}", fieldDefinition.InternalName, ex.Message);
                }
                finally
                {
                    returnField = hostWeb.Fields.GetByInternalNameOrTitle(fieldDefinition.InternalName);
                    hostWeb.Context.Load(returnField, fd => fd.Id, fd => fd.Title, fd => fd.Indexed, fd => fd.InternalName, fd => fd.CanBeDeleted, fd => fd.Required);
                    hostWeb.Context.ExecuteQuery();
                }
            }

            return returnField;
        }

        /// <summary>
        /// Provision a list to the specified web using the list definition
        /// </summary>
        /// <param name="web">Client Context web</param>
        /// <param name="listDefinition">Hydrated list definition from JSON or Object</param>
        /// <param name="provisionerChoices">(OPTIONAL) deserialized choices from JSON</param>
        public static List CreateListFromDefinition(this Web web, SPListDefinition listDefinition, List<SiteProvisionerFieldChoiceModel> provisionerChoices = null)
        {
            var webContext = web.Context;

            var siteColumns = new List<SPFieldDefinitionModel>();
            var afterProvisionChanges = false;

            // Content Type
            var listName = listDefinition.ListName;
            var listDescription = listDefinition.ListDescription;


            // check to see if Picture library named Photos already exists
            ListCollection allLists = web.Lists;
            IEnumerable<List> foundLists = webContext.LoadQuery(allLists.Where(list => list.Title == listName)
                .Include(arl => arl.Title, arl => arl.Id, arl => arl.ContentTypes, ol => ol.RootFolder, ol => ol.EnableVersioning, ol => ol.EnableFolderCreation, ol => ol.ContentTypesEnabled));
            webContext.ExecuteQueryRetry();

            List listToProvision = foundLists.FirstOrDefault();
            if (listToProvision == null)
            {
                ListCreationInformation listToProvisionInfo = listDefinition.ToCreationObject();
                listToProvision = web.Lists.Add(listToProvisionInfo);
                webContext.Load(listToProvision, arl => arl.Title, arl => arl.Id, arl => arl.ContentTypes, ol => ol.RootFolder, ol => ol.EnableVersioning, ol => ol.EnableFolderCreation, ol => ol.ContentTypesEnabled);
                webContext.ExecuteQuery();
            }

            if (listDefinition.Versioning && !listToProvision.EnableVersioning)
            {
                afterProvisionChanges = true;
                listToProvision.EnableVersioning = true;
                if (listDefinition.ListTemplate == ListTemplateType.DocumentLibrary)
                {
                    listToProvision.EnableMinorVersions = true;
                }
            }

            if (listDefinition.ContentTypeEnabled && !listToProvision.ContentTypesEnabled)
            {
                afterProvisionChanges = true;
                listToProvision.ContentTypesEnabled = true;
            }

            if (listDefinition.EnableFolderCreation && !listToProvision.EnableFolderCreation)
            {
                afterProvisionChanges = true;
                listToProvision.EnableFolderCreation = true;
            }

            if (afterProvisionChanges)
            {
                listToProvision.Update();
                webContext.Load(listToProvision);
                webContext.ExecuteQueryRetry();
            }

            webContext.Load(listToProvision, arl => arl.Title, arl => arl.Id, arl => arl.ContentTypes, ol => ol.RootFolder, ol => ol.EnableVersioning, ol => ol.EnableFolderCreation, ol => ol.ContentTypesEnabled);
            webContext.ExecuteQueryRetry();


            return listToProvision;
        }

        /// <summary>
        /// Generate a portable JSON object from the List Template
        /// </summary>
        /// <param name="context">Client Context web</param>
        /// <param name="list">Hydrated SharePoint list Object</param>
        /// <param name="ExpandObjects">true - enumerate fields, views, content types</param>
        /// <param name="logger">Logger implementation for Verbose/Exception handling</param>
        /// <param name="skiptypes">Collection of field types to be used as a filter statement</param>
        /// <param name="siteGroups">Collection of hostWeb groups</param>
        /// <returns></returns>
        public static SPListDefinition GetListDefinition(this ClientContext context, List list, bool ExpandObjects, ITraceLogger logger, IEnumerable<FieldType> skiptypes, IEnumerable<Microsoft.SharePoint.Client.Group> siteGroups = null)
        {
            logger.LogInformation("Processing Client Context for list {0}", list.Title);

            var listDefinition = context.GetListDefinition(context.Web, list, ExpandObjects, logger, skiptypes, siteGroups);
            return listDefinition;
        }

        /// <summary>
        /// Generate a portable JSON object from the List Template
        /// </summary>
        /// <param name="context">Client Context web</param>
        /// <param name="hostWeb">Client Context web</param>
        /// <param name="list">Hydrated SharePoint list Object</param>
        /// <param name="ExpandObjects">true - enumerate fields, views, content types</param>
        /// <param name="logger">Logger implementation for Verbose/Exception handling</param>
        /// <param name="skiptypes">Collection of field types to be used as a filter statement</param>
        /// <param name="siteGroups">Collection of hostWeb groups</param>
        /// <returns></returns>
        public static SPListDefinition GetListDefinition(this ClientContext context, Web hostWeb, List list, bool ExpandObjects, ITraceLogger logger, IEnumerable<FieldType> skiptypes, IEnumerable<Microsoft.SharePoint.Client.Group> siteGroups = null)
        {
            logger.LogInformation("Processing Web list {0}", list.Title);

            if (!hostWeb.IsPropertyAvailable(ctx => ctx.ServerRelativeUrl))
            {
                hostWeb.Context.Load(hostWeb, ctx => ctx.ServerRelativeUrl);
                hostWeb.Context.ExecuteQueryRetry();
            }

            list.EnsureProperties(
                    lctx => lctx.Id,
                    lctx => lctx.Title,
                    lctx => lctx.Description,
                    lctx => lctx.DefaultViewUrl,
                    lctx => lctx.OnQuickLaunch,
                    lctx => lctx.BaseTemplate,
                    lctx => lctx.BaseType,
                    lctx => lctx.CrawlNonDefaultViews,
                    lctx => lctx.Created,
                    lctx => lctx.ContentTypesEnabled,
                    lctx => lctx.CreatablesInfo,
                    lctx => lctx.EnableFolderCreation,
                    lctx => lctx.EnableModeration,
                    lctx => lctx.EnableVersioning,
                    lctx => lctx.Hidden,
                    lctx => lctx.IsApplicationList,
                    lctx => lctx.IsCatalog,
                    lctx => lctx.IsSiteAssetsLibrary,
                    lctx => lctx.IsPrivate,
                    lctx => lctx.IsSystemList,
                    lctx => lctx.RootFolder.ServerRelativeUrl,
                    lctx => lctx.SchemaXml,
                    lctx => lctx.LastItemModifiedDate,
                    lctx => lctx.LastItemUserModifiedDate,
                    lctx => lctx.ListExperienceOptions,
                    lctx => lctx.TemplateFeatureId);

            var weburl = TokenHelper.EnsureTrailingSlash(hostWeb.ServerRelativeUrl);
            var listTemplateType = list.GetListTemplateType();

            var listdefinition = new SPListDefinition()
            {
                Id = list.Id,
                ListName = list.Title,
                ListDescription = list.Description,
                ServerRelativeUrl = list.DefaultViewUrl,
                BaseTemplate = list.BaseTemplate,
                ListTemplate = listTemplateType,
                Created = list.Created,
                LastItemModifiedDate = list.LastItemModifiedDate,
                LastItemUserModifiedDate = list.LastItemUserModifiedDate,
                QuickLaunch = list.OnQuickLaunch ? QuickLaunchOptions.On : QuickLaunchOptions.Off,
                ContentTypeEnabled = list.ContentTypesEnabled,
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
                var contentTypesFieldset = new List<dynamic>();
                var definitionListFields = new List<SPFieldDefinitionModel>();
                var listurl = TokenHelper.EnsureTrailingSlash(list.RootFolder.ServerRelativeUrl);

                // content types
                var listContentType = list.Context.LoadQuery(list.ContentTypes
                    .Include(
                        ict => ict.Id,
                        ict => ict.Group,
                        ict => ict.Description,
                        ict => ict.Name,
                        ict => ict.Hidden,
                        ict => ict.JSLink,
                        ict => ict.FieldLinks,
                        ict => ict.Fields));
                // list fields
                var listFields = list.Context.LoadQuery(list.Fields.Where(wf => wf.ReadOnlyField == false && wf.Hidden == false)
                    .Include(
                       fctx => fctx.Id,
                       fctx => fctx.AutoIndexed,
                       fctx => fctx.CanBeDeleted,
                       fctx => fctx.DefaultFormula,
                       fctx => fctx.DefaultValue,
                       fctx => fctx.Group,
                       fctx => fctx.Description,
                       fctx => fctx.EnforceUniqueValues,
                       fctx => fctx.FieldTypeKind,
                       fctx => fctx.Filterable,
                       fctx => fctx.FromBaseType,
                       fctx => fctx.Hidden,
                       fctx => fctx.Indexed,
                       fctx => fctx.InternalName,
                       fctx => fctx.JSLink,
                       fctx => fctx.NoCrawl,
                       fctx => fctx.ReadOnlyField,
                       fctx => fctx.Required,
                       fctx => fctx.SchemaXml,
                       fctx => fctx.Scope,
                       fctx => fctx.Title
                       ));
                // list views
                var listViews = list.Context.LoadQuery(list.Views
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

                list.Context.ExecuteQueryRetry();


                if (listContentType != null && listContentType.Any())
                {
                    listdefinition.ContentTypes = new List<SPContentTypeDefinition>();
                    foreach (var contenttype in listContentType)
                    {
                        logger.LogInformation("Processing list {0} content type {1}", list.Title, contenttype.Name);

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
                            foreach (var cfieldlink in contenttype.FieldLinks)
                            {
                                ctypemodel.FieldLinks.Add(new SPFieldLinkDefinitionModel()
                                {
                                    Id = cfieldlink.Id,
                                    Name = cfieldlink.Name,
                                    Hidden = cfieldlink.Hidden,
                                    Required = cfieldlink.Required
                                });

                                contentTypesFieldset.Add(new { ctypeid = contenttype.Id.StringValue, name = cfieldlink.Name });
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


                if (listFields != null && listFields.Any())
                {
                    var filteredListFields = listFields.Where(lf => !skiptypes.Any(st => lf.FieldTypeKind == st)).ToList();
                    logger.LogWarning("Processing list {0} found {1} fields to be processed", list.Title, filteredListFields.Count());

                    foreach (Field listField in listFields)
                    {
                        logger.LogInformation("Processing list {0} field {1}", list.Title, listField.InternalName);

                        try
                        {
                            var fieldXml = listField.SchemaXml;
                            if (!string.IsNullOrEmpty(fieldXml))
                            {
                                var xdoc = XDocument.Parse(fieldXml, LoadOptions.PreserveWhitespace);
                                var xField = xdoc.Element("Field");
                                var xSourceID = xField.Attribute("SourceID");
                                //if (xSourceID != null && xSourceID.Value.IndexOf(ConstantsXmlNamespaces.SharePointNS.NamespaceName, StringComparison.CurrentCultureIgnoreCase) < 0)
                                //{
                                //    continue; // skip processing an OOTB field
                                //}
                                var customField = context.RetrieveField(listField, logger, siteGroups, xField);
                                if (xSourceID != null)
                                {
                                    customField.SourceID = xSourceID.Value;
                                }
                                definitionListFields.Add(customField);

                                if (customField.FieldTypeKind == FieldType.Lookup)
                                {
                                    listdefinition.ListDependency.Add(customField.LookupListName);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            logger.LogError(ex, "Failed to parse field {0} MSG:{1}", listField.InternalName, ex.Message);
                        }
                    }

                    listdefinition.FieldDefinitions = definitionListFields;
                }


                if (listViews != null && listViews.Any())
                {
                    listdefinition.InternalViews = new List<SPViewDefinitionModel>();
                    listdefinition.Views = new List<SPViewDefinitionModel>();

                    foreach (var view in listViews)
                    {
                        logger.LogInformation("Processing list {0} view {1}", list.Title, view.Title);

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
            }

            return listdefinition;
        }

        /// <summary>
        /// Add web part to a wiki style page
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="folder">System name of the wiki page library - typically sitepages</param>
        /// <param name="webPart">Information about the web part to insert</param>
        /// <param name="page">Page to add the web part on</param>
        /// <param name="row">Row of the wiki table that should hold the inserted web part</param>
        /// <param name="col">Column of the wiki table that should hold the inserted web part</param>
        /// <param name="addSpace">Does a blank line need to be added after the web part (to space web parts)</param>
        /// <exception cref="System.ArgumentException">Thrown when folder or page is a zero-length string or contains only white space</exception>
        /// <exception cref="System.ArgumentNullException">Thrown when folder, webPart or page is null</exception>
        public static void AddWebPartToWikiPage(this Web web, string folder, WebPartEntity webPart, string page, int row, int col, bool addSpace, bool something)
        {
            if (string.IsNullOrEmpty(folder))
            {
                throw (folder == null)
                  ? new ArgumentNullException("folder")
                  : new ArgumentException("Empty string for folder", "folder");
            }

            if (webPart == null)
            {
                throw new ArgumentNullException("webPart");
            }

            if (string.IsNullOrEmpty(page))
            {
                throw (page == null)
                  ? new ArgumentNullException("page")
                  : new ArgumentException("Empty string for page", "page");
            }

            if (!web.IsObjectPropertyInstantiated("ServerRelativeUrl"))
            {
                web.Context.Load(web, w => w.ServerRelativeUrl);
                web.Context.ExecuteQueryRetry();
            }

            var webServerRelativeUrl = UrlUtility.EnsureTrailingSlash(web.ServerRelativeUrl);
            var serverRelativeUrl = UrlUtility.Combine(folder, page);
            AddWebPartToWikiPage(web, webServerRelativeUrl + serverRelativeUrl, webPart, row, col, addSpace);
        }

        /// <summary>
        /// Add web part to a wiki style page
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="serverRelativePageUrl">Server relative url of the page to add the webpart to</param>
        /// <param name="webPart">Information about the web part to insert</param>
        /// <param name="row">Row of the wiki table that should hold the inserted web part</param>
        /// <param name="col">Column of the wiki table that should hold the inserted web part</param>
        /// <param name="addSpace">Does a blank line need to be added after the web part (to space web parts)</param>
        /// <exception cref="System.ArgumentException">Thrown when serverRelativePageUrl is a zero-length string or contains only white space</exception>
        /// <exception cref="System.ArgumentNullException">Thrown when serverRelativePageUrl or webPart is null</exception>
        public static void AddWebPartToWikiPage(this Web web, string serverRelativePageUrl, WebPartEntity webPart, int row, int col, bool addSpace)
        {
            if (string.IsNullOrEmpty(serverRelativePageUrl))
            {
                throw (serverRelativePageUrl == null)
                  ? new ArgumentNullException("serverRelativePageUrl")
                  : new ArgumentException("Empty parameter", "serverRelativePageUrl");
            }

            if (webPart == null)
            {
                throw new ArgumentNullException("webPart");
            }

            File webPartPage = web.GetFileByServerRelativeUrl(serverRelativePageUrl);

            if (webPartPage == null)
            {
                return;
            }

            web.Context.Load(webPartPage, wp => wp.ListItemAllFields);
            web.Context.ExecuteQueryRetry();

            string wikiField = (string)webPartPage.ListItemAllFields["WikiField"];

            LimitedWebPartManager limitedWebPartManager = webPartPage.GetLimitedWebPartManager(PersonalizationScope.Shared);
            WebPartDefinition oWebPartDefinition = limitedWebPartManager.ImportWebPart(webPart.WebPartXml);
            WebPartDefinition wpdNew = limitedWebPartManager.AddWebPart(oWebPartDefinition.WebPart, "wpz", 0);
            web.Context.Load(wpdNew);
            web.Context.ExecuteQueryRetry();

            // Close all BR tags
            Regex brRegex = new Regex("<br>", RegexOptions.IgnoreCase);

            wikiField = brRegex.Replace(wikiField, "<br/>");

            XmlDocument xd = new XmlDocument();
            xd.PreserveWhitespace = true;
            xd.LoadXml(wikiField);

            // Sometimes the wikifield content seems to be surrounded by an additional div? 
            XmlElement layoutsTable = xd.SelectSingleNode("div/div/table") as XmlElement;
            if (layoutsTable == null)
            {
                layoutsTable = xd.SelectSingleNode("div/table") as XmlElement;
            }

            XmlElement layoutsZoneInner = layoutsTable.SelectSingleNode(string.Format("tbody/tr[{0}]/td[{1}]/div/div", row, col)) as XmlElement;
            // - space element
            XmlElement space = xd.CreateElement("p");
            XmlText text = xd.CreateTextNode(" ");
            space.AppendChild(text);

            // - wpBoxDiv
            XmlElement wpBoxDiv = xd.CreateElement("div");
            layoutsZoneInner.AppendChild(wpBoxDiv);

            if (addSpace)
            {
                layoutsZoneInner.AppendChild(space);
            }

            XmlAttribute attribute = xd.CreateAttribute("class");
            wpBoxDiv.Attributes.Append(attribute);
            attribute.Value = "ms-rtestate-read ms-rte-wpbox";
            attribute = xd.CreateAttribute("contentEditable");
            wpBoxDiv.Attributes.Append(attribute);
            attribute.Value = "false";
            // - div1
            XmlElement div1 = xd.CreateElement("div");
            wpBoxDiv.AppendChild(div1);
            div1.IsEmpty = false;
            attribute = xd.CreateAttribute("class");
            div1.Attributes.Append(attribute);
            attribute.Value = "ms-rtestate-read " + wpdNew.Id.ToString("D");
            attribute = xd.CreateAttribute("id");
            div1.Attributes.Append(attribute);
            attribute.Value = "div_" + wpdNew.Id.ToString("D");
            // - div2
            XmlElement div2 = xd.CreateElement("div");
            wpBoxDiv.AppendChild(div2);
            div2.IsEmpty = false;
            attribute = xd.CreateAttribute("style");
            div2.Attributes.Append(attribute);
            attribute.Value = "display:none";
            attribute = xd.CreateAttribute("id");
            div2.Attributes.Append(attribute);
            attribute.Value = "vid_" + wpdNew.Id.ToString("D");

            ListItem listItem = webPartPage.ListItemAllFields;
            listItem["WikiField"] = xd.OuterXml;
            listItem.Update();
            web.Context.ExecuteQueryRetry();

        }

        /// <summary>
        /// Adds or Updates an existing Custom Action [ScriptSrc] into the [Web] Custom Actions
        /// </summary>
        /// <param name="web"></param>
        /// <param name="customactionname"></param>
        /// <param name="customactionurl"></param>
        /// <param name="sequence"></param>
        public static void AddOrUpdateCustomActionLink(this Web web, string customactionname, string customactionurl, int sequence)
        {
            var sitecustomActions = web.GetCustomActions();
            UserCustomAction cssAction = null;
            if (web.CustomActionExists(customactionname))
            {
                cssAction = sitecustomActions.FirstOrDefault(fod => fod.Name == customactionname);
            }
            else
            {
                // Build a custom action to write a link to our new CSS file
                cssAction = web.UserCustomActions.Add();
                cssAction.Name = customactionname;
                cssAction.Location = "ScriptLink";
            }

            cssAction.Sequence = sequence;
            cssAction.ScriptSrc = customactionurl;
            cssAction.Update();
            web.Context.ExecuteQueryRetry();
        }

        /// <summary>
        /// Adds or Updates an existing Custom Action [ScriptBlock] into the [Web] Custom Actions
        /// </summary>
        /// <param name="web"></param>
        /// <param name="customactionname"></param>
        /// <param name="customActionBlock"></param>
        /// <param name="sequence"></param>
        public static void AddOrUpdateCustomActionLinkBlock(this Web web, string customactionname, string customActionBlock, int sequence)
        {
            var sitecustomActions = web.GetCustomActions();
            UserCustomAction cssAction = null;
            if (web.CustomActionExists(customactionname))
            {
                cssAction = sitecustomActions.FirstOrDefault(fod => fod.Name == customactionname);
            }
            else
            {
                // Build a custom action to write a link to our new CSS file
                cssAction = web.UserCustomActions.Add();
                cssAction.Name = customactionname;
                cssAction.Location = "ScriptLink";
            }

            cssAction.Sequence = sequence;
            cssAction.ScriptBlock = customActionBlock;
            cssAction.Update();
            web.Context.ExecuteQueryRetry();
        }

        /// <summary>
        /// Will remove the custom action if one exists
        /// </summary>
        /// <param name="web"></param>
        /// <param name="customactionname"></param>
        public static bool RemoveCustomActionLink(this Web web, string customactionname)
        {
            if (web.CustomActionExists(customactionname))
            {
                var cssAction = web.GetCustomActions().FirstOrDefault(fod => fod.Name == customactionname || fod.Title == customactionname);
                web.DeleteCustomAction(cssAction.Id);
                return true;
            }
            return false;
        }
    }
}
