using IaC.Core.Models;
using IaC.Core.Extensions;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WebParts;
using OfficeDevPnP.Core.Entities;
using OfficeDevPnP.Core.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;

namespace IaC.Powershell.Extensions
{
    public static class WebExtensions
    {
        /// <summary>
        /// Provision a list to the specified web using the list definition
        /// </summary>
        /// <param name="web">Client Context web</param>
        /// <param name="listDef">Hydrated list definition from JSON or Object</param>
        /// <param name="loggerVerbose">TODO: convert to static logger</param>
        /// <param name="loggerWarning">TODO: convert to static logger</param>
        /// <param name="loggerError">TODO: convert to static logger</param>
        /// <param name="SiteGroups">Collection of provisioned SharePoint group for field definitions</param>
        /// <param name="JsonFilePath">(OPTIONAL) file path to JSON folder</param>
        public static void CreateListFromDefinition(this Web web, SPListDefinition listDef, Action<string, string[]> loggerVerbose, Action<string, string[]> loggerWarning, Action<Exception, string, string[]> loggerError, List<SPGroupDefinitionModel> SiteGroups = null, string JsonFilePath = null)
        {
            var webContext = web.Context;

            var siteColumns = new List<SPFieldDefinitionModel>();
            var afterProvisionChanges = false;

            // Content Type
            var listName = listDef.ListName;
            var listDescription = listDef.ListDescription;


            // check to see if Picture library named Photos already exists
            ListCollection allLists = web.Lists;
            IEnumerable<List> foundLists = webContext.LoadQuery(allLists.Where(list => list.Title == listName)
                .Include(arl => arl.Title, arl => arl.Id, arl => arl.ContentTypes, ol => ol.RootFolder, ol => ol.EnableVersioning, ol => ol.EnableFolderCreation, ol => ol.ContentTypesEnabled));
            webContext.ExecuteQueryRetry();

            List listToProvision = foundLists.FirstOrDefault();
            if (listToProvision == null)
            {
                ListCreationInformation listToProvisionInfo = listDef.ToCreationObject();
                listToProvision = web.Lists.Add(listToProvisionInfo);
                webContext.Load(listToProvision, arl => arl.Title, arl => arl.Id, arl => arl.ContentTypes, ol => ol.RootFolder, ol => ol.EnableVersioning, ol => ol.EnableFolderCreation, ol => ol.ContentTypesEnabled);
                webContext.ExecuteQuery();
            }

            if (listDef.Versioning && !listToProvision.EnableVersioning)
            {
                afterProvisionChanges = true;
                listToProvision.EnableVersioning = true;
                if (listDef.ListTemplate == ListTemplateType.DocumentLibrary)
                {
                    listToProvision.EnableMinorVersions = true;
                }
            }

            if (listDef.ContentTypeEnabled && !listToProvision.ContentTypesEnabled)
            {
                afterProvisionChanges = true;
                listToProvision.ContentTypesEnabled = true;
            }

            if (listDef.EnableFolderCreation && !listToProvision.EnableFolderCreation)
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

            if (listDef.ContentTypeEnabled && listDef.HasContentTypes)
            {
                foreach (var contentDef in listDef.ContentTypes)
                {
                    if (!listToProvision.ContentTypeExistsByName(contentDef.ContentTypeName))
                    {
                        var ctypeInfo = contentDef.ToCreationObject();
                        var accessContentType = listToProvision.ContentTypes.Add(ctypeInfo);
                        listToProvision.Update();
                        webContext.Load(accessContentType, tycp => tycp.Id, tycp => tycp.Name);
                        webContext.ExecuteQueryRetry();

                        if (contentDef.DefaultContentType)
                        {
                            listToProvision.SetDefaultContentTypeToList(accessContentType);
                        }
                    }
                }
            }


            // Site Columns
            foreach (var fieldDef in listDef.FieldDefinitions)
            {
                var column = listToProvision.CreateColumn(fieldDef, loggerVerbose, loggerError, SiteGroups, JsonFilePath);
                if (column == null)
                {
                    loggerWarning("Failed to create column {0}.", new string[] { fieldDef.InternalName });
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

            if (listDef.ContentTypeEnabled && listDef.HasContentTypes)
            {
                foreach (var contentDef in listDef.ContentTypes)
                {
                    var contentTypeName = contentDef.ContentTypeName;
                    var accessContentTypes = listToProvision.ContentTypes;
                    IEnumerable<ContentType> allContentTypes = webContext.LoadQuery(accessContentTypes.Where(f => f.Name == contentTypeName).Include(tcyp => tcyp.Id, tcyp => tcyp.Name));
                    webContext.ExecuteQueryRetry();

                    if (allContentTypes != null)
                    {
                        var accessContentType = allContentTypes.FirstOrDefault();
                        foreach (var fieldInternalName in contentDef.FieldLinkRefs)
                        {
                            var column = siteColumns.FirstOrDefault(f => f.InternalName == fieldInternalName);
                            if (column != null)
                            {
                                var fieldLinks = accessContentType.FieldLinks;
                                webContext.Load(fieldLinks, cf => cf.Include(inc => inc.Id, inc => inc.Name));
                                webContext.ExecuteQueryRetry();

                                var convertedInternalName = column.DisplayNameMasked;
                                if (!fieldLinks.Any(a => a.Name == column.InternalName || a.Name == convertedInternalName))
                                {
                                    loggerVerbose("Content Type {0} Adding Field {1}", new string[] { contentTypeName, column.InternalName });
                                    var siteColumn = listToProvision.GetFieldById<Field>(column.FieldGuid);
                                    webContext.ExecuteQueryRetry();

                                    var flink = new FieldLinkCreationInformation();
                                    flink.Field = siteColumn;
                                    var flinkstub = accessContentType.FieldLinks.Add(flink);
                                    //if(fieldDef.Required) flinkstub.Required = fieldDef.Required;
                                    accessContentType.Update(false);
                                    webContext.ExecuteQueryRetry();
                                }
                            }
                            else
                            {
                                loggerWarning("Failed to create column {0}.", new string[] { fieldInternalName });
                            }
                        }
                    }
                }
            }


            // Views
            if (listDef.Views != null && listDef.Views.Count() > 0)
            {
                ViewCollection views = listToProvision.Views;
                webContext.Load(views, f => f.Include(inc => inc.Id, inc => inc.Hidden, inc => inc.Title, inc => inc.DefaultView));
                webContext.ExecuteQueryRetry();

                foreach (var viewDef in listDef.Views)
                {
                    try
                    {
                        if (views.Any(v => v.Title.Equals(viewDef.Title, StringComparison.CurrentCultureIgnoreCase)))
                        {
                            loggerVerbose("View {0} found in list {1}", new string[] { viewDef.Title, listName });
                            continue;
                        }

                        var view = listToProvision.CreateView(viewDef.InternalName, viewDef.ViewCamlType, viewDef.FieldRefName, viewDef.RowLimit, viewDef.DefaultView, viewDef.QueryXml, viewDef.PersonalView, viewDef.PagedView);
                        webContext.Load(view, v => v.Title, v => v.Id, v => v.ServerRelativeUrl);
                        webContext.ExecuteQueryRetry();

                        view.Title = viewDef.Title;
                        if (viewDef.HasJsLink)
                        {
                            view.JSLink = viewDef.JsLink;
                        }
                        view.Update();
                        webContext.ExecuteQueryRetry();
                    }
                    catch (Exception ex)
                    {
                        loggerError(ex, "Failed to create view {0} with XML:{1}", new string[] { viewDef.Title, viewDef.QueryXml });
                    }
                }
            }

            // List Data upload
            if (listDef.ListItems != null && listDef.ListItems.Count() > 0)
            {
                foreach (var listItem in listDef.ListItems)
                {

                    // Process the record into the notification email list
                    var itemCreateInfo = new ListItemCreationInformation();

                    var newSPListItem = listToProvision.AddItem(itemCreateInfo);
                    newSPListItem["Title"] = listItem.Title;
                    foreach (var listItemData in listItem.ColumnValues)
                    {
                        newSPListItem[listItemData.FieldName] = listItemData.FieldValue;
                    }

                    newSPListItem.Update();
                    webContext.ExecuteQueryRetry();
                }
            }
        }
    }
}
