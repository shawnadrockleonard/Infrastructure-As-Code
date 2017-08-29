using InfrastructureAsCode.Core.Models;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using OfficeDevPnP.Core;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Entities;
using OfficeDevPnP.Core.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace InfrastructureAsCode.Core.Extensions
{
    public static partial class ListExtensions
    {
        /// <summary>
        /// Adds a field to a list
        /// </summary>
        /// <param name="list">List to process</param>
        /// <param name="fieldAsXml">The XML declaration of SiteColumn definition</param>
        /// <param name="AddToDefaultView">Optionally add to the default view</param>
        /// <param name="executeQuery">Optionally skip the executeQuery action</param>
        /// <returns>The newly created field or existing field.</returns>
        public static Field CreateField(this List list, string fieldAsXml, bool AddToDefaultView = false, bool executeQuery = true)
        {
            var fields = list.Fields;
            list.Context.Load(fields);
            list.Context.ExecuteQueryRetry();

            var xd = XDocument.Parse(fieldAsXml);
            if (xd.Root != null)
            {
                var ns = xd.Root.Name.Namespace;

                var fieldNode = (from f in xd.Elements(ns + "Field") select f).FirstOrDefault();

                if (fieldNode != null)
                {
                    string id = string.Empty;
                    if (fieldNode.Attribute("ID") != null)
                    {
                        id = fieldNode.Attribute("ID").Value;
                    }
                    else
                    {
                        id = "<No ID specified in XML>";
                    }
                    var name = fieldNode.Attribute("Name").Value;

                    Log.Info("FieldAndContentTypeExtensions", "CreateField {0} with ID {1}", name, id);
                }
            }
            var field = fields.AddFieldAsXml(fieldAsXml, AddToDefaultView, AddFieldOptions.AddFieldInternalNameHint);
            list.Update();

            if (executeQuery)
            {
                list.Context.ExecuteQueryRetry();
            }

            return field;
        }

        /// <summary>
        /// Provisions a column based on the field defintion to the host list
        /// </summary>
        /// <param name="hostList">The instantiated list/library to which the field will be added</param>
        /// <param name="fieldDefinition">The definition for the field</param>
        /// <param name="loggerVerbose">Provides a method for verbose logging</param>
        /// <param name="loggerError">Provides a method for exception logging</param>
        /// <param name="SiteGroups">(OPTIONAL) collection of group, required if this is a PeoplePicker field</param>
        /// <param name="provisionerChoices">(OPTIONAL) deserialized choices from JSON</param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException">For field definitions that do not contain all required data</exception>
        public static Field CreateListColumn(this List hostList, SPFieldDefinitionModel fieldDefinition, Action<string, string[]> loggerVerbose, Action<string, string[]> loggerError, List<SPGroupDefinitionModel> SiteGroups, List<SiteProvisionerFieldChoiceModel> provisionerChoices = null)
        {

            if (fieldDefinition == null)
            {
                throw new ArgumentNullException("fieldDef", "Field definition is required.");
            }

            if (fieldDefinition.LoadFromJSON && (provisionerChoices == null || !provisionerChoices.Any(pc => pc.FieldInternalName == fieldDefinition.InternalName)))
            {
                throw new ArgumentNullException("provisionerChoices", string.Format("You must specify a collection of field choices for the field {0}", fieldDefinition.Title));
            }

            var fields = hostList.Fields;
            hostList.Context.Load(fields, fc => fc.Include(f => f.Id, f => f.InternalName, f => f.Title, f => f.JSLink, f => f.Indexed, f => f.CanBeDeleted, f => f.Required));
            hostList.Context.ExecuteQueryRetry();

            var returnField = fields.FirstOrDefault(f => f.Id == fieldDefinition.FieldGuid || f.InternalName == fieldDefinition.InternalName);
            if (returnField == null)
            {
                try
                {
                    var baseFieldXml = hostList.CreateFieldDefinition(fieldDefinition, SiteGroups, provisionerChoices);
                    loggerVerbose("Provision field {0} with XML:{1}", new string[] { fieldDefinition.InternalName, baseFieldXml });

                    // Should throw an exception if the field ID or Name exist in the list
                    var baseField = hostList.CreateField(baseFieldXml, fieldDefinition.AddToDefaultView, executeQuery: false);
                    hostList.Context.ExecuteQueryRetry();
                }
                catch (Exception ex)
                {
                    var msg = ex.Message;
                    loggerError("EXCEPTION: field {0} with message {1}", new string[] { fieldDefinition.InternalName, msg });
                }
                finally
                {
                    returnField = hostList.Fields.GetByInternalNameOrTitle(fieldDefinition.InternalName);
                    hostList.Context.Load(returnField, fd => fd.Id, fd => fd.Title, fd => fd.Indexed, fd => fd.InternalName, fd => fd.CanBeDeleted, fd => fd.Required);
                    hostList.Context.ExecuteQueryRetry();
                }
            }

            return returnField;
        }

        /// <summary>
        /// Retrieves or creates the folder as a ListItem
        /// </summary>
        /// <param name="onlineLibrary"></param>
        /// <param name="destinationFolder"></param>
        /// <param name="folderName"></param>
        /// <param name="defaultLastItemId"></param>
        /// <returns>The listitem as a folder</returns>
        public static Folder GetOrCreateFolder(this List onlineLibrary, Folder destinationFolder, string folderName, int? defaultLastItemId = default(int?))
        {
            destinationFolder.EnsureProperties(afold => afold.ServerRelativeUrl, afold => afold.Folders);
            var folderRelativeUrl = destinationFolder.ServerRelativeUrl;
            // Remove invalid characters
            var trimmedFolder = HelperExtensions.GetCleanDirectory(folderName, string.Empty);
            ListItem folderItem = null;

            var camlFields = new string[] { "Title", "ContentType", "ID" };
            var camlViewFields = CAML.ViewFields(camlFields.Select(s => CAML.FieldRef(s)).ToArray());


            var camlClause = CAML.And(
                        CAML.And(
                            CAML.Eq(CAML.FieldValue("FileDirRef", FieldType.Text.ToString("f"), folderRelativeUrl)),
                            CAML.Or(
                                CAML.Eq(CAML.FieldValue("LinkFilename", FieldType.Text.ToString("f"), trimmedFolder)),
                                CAML.Eq(CAML.FieldValue("Title", FieldType.Text.ToString("f"), trimmedFolder))
                            )
                        )
                        ,
                    CAML.Eq(CAML.FieldValue("FSObjType", FieldType.Integer.ToString("f"), 1.ToString()))
                    );

            var camlQueries = onlineLibrary.SafeCamlClauseFromThreshold(camlClause, defaultLastItemId);
            foreach (var camlAndValue in camlQueries)
            {
                var camlWhereClause = CAML.Where(camlAndValue);
                var camlQuery = new CamlQuery()
                {
                    ViewXml = CAML.ViewQuery(ViewScope.RecursiveAll, camlWhereClause, string.Empty, camlViewFields, 5)
                };
                var listItems = onlineLibrary.GetItems(camlQuery);
                onlineLibrary.Context.Load(listItems);
                onlineLibrary.Context.ExecuteQueryRetry();

                if (listItems.Count() > 0)
                {
                    folderItem = listItems.FirstOrDefault();
                    System.Diagnostics.Trace.TraceInformation("Item {0} exists in the destination folder.  Skip item creation file.....", folderName);
                    break;
                }
            };

            if (folderItem != null)
            {
                return folderItem.Folder;
            }

            try
            {
                var info = new ListItemCreationInformation();
                info.UnderlyingObjectType = FileSystemObjectType.Folder;
                info.LeafName = trimmedFolder;
                info.FolderUrl = folderRelativeUrl;

                folderItem = onlineLibrary.AddItem(info);
                folderItem["Title"] = trimmedFolder;
                folderItem.Update();
                onlineLibrary.Context.ExecuteQueryRetry();
                System.Diagnostics.Trace.TraceInformation("{0} folder Created", trimmedFolder);
                return folderItem.Folder;
            }
            catch (Exception Ex)
            {
                System.Diagnostics.Trace.TraceError("Failed to create or get folder for name {0} MSG:{1}", folderName, Ex.Message);
            }

            return null;
        }

        /// <summary>
        /// Will retreive the folder or create if it does not exist
        /// </summary>
        /// <param name="destinationFolder"></param>
        /// <param name="folderName"></param>
        /// <returns></returns>
        public static Folder GetOrCreateFolder(this Folder destinationFolder, string folderName)
        {
            // clean the folder name
            var trimmedFolder = folderName.Trim().Replace("_", " ");

            // setup processing of folder in the parent folder
            var currentFolder = destinationFolder;
            destinationFolder.Context.Load(destinationFolder, pf => pf.Name, pf => pf.Folders);
            destinationFolder.Context.ExecuteQuery();

            if (!destinationFolder.FolderExists(trimmedFolder))
            {
                currentFolder = destinationFolder.EnsureFolder(trimmedFolder);
                //this.ClientContext.Load(curFolder);
                destinationFolder.Context.ExecuteQueryRetry();
                System.Diagnostics.Trace.TraceInformation(".......... successfully created folder {0}....", trimmedFolder);
            }
            else
            {
                currentFolder = destinationFolder.Folders.FirstOrDefault(f => f.Name == trimmedFolder);
                System.Diagnostics.Trace.TraceInformation(".......... reading folder {0}....", trimmedFolder);
            }

            return currentFolder;
        }

        /// <summary>
        /// Build folder path from root of the parent list
        /// </summary>
        /// <param name="parentList"></param>
        /// <param name="folderUrl"></param>
        /// <returns></returns>
        public static Folder ListEnsureFolder(this List parentList, string folderUrl)
        {
            if (!parentList.IsPropertyAvailable("RootFolder"))
            {
                parentList.EnsureProperties(pl => pl.RootFolder, pl => pl.RootFolder.ServerRelativeUrl);
            }

            var listUri = new Uri(parentList.RootFolder.ServerRelativeUrl);
            var relativeUri = listUri.MakeRelativeUri(new Uri(folderUrl));
            var relativeUrl = folderUrl.Replace(parentList.RootFolder.ServerRelativeUrl, "");

            var folder = parentList.RootFolder.ListEnsureFolder(folderUrl);
            return folder;
        }

        /// <summary>
        /// Build folder path
        /// </summary>
        /// <param name="parentFolder"></param>
        /// <param name="folderUrl"></param>
        /// <returns></returns>
        public static Folder ListEnsureFolder(this Folder parentFolder, string folderUrl)
        {
            var folderNames = folderUrl.Split(new string[] { "/", "\\" }, StringSplitOptions.RemoveEmptyEntries);
            var folderName = folderNames[0];

            var ctx = parentFolder.Context;
            if (!parentFolder.IsPropertyAvailable("Folders"))
            {
                ctx.Load(parentFolder, inn => inn.ServerRelativeUrl, inn => inn.Folders);
                ctx.ExecuteQueryRetry();
            }

            var folder = parentFolder.EnsureFolder(folderName);

            if (folderNames.Length > 1)
            {
                var subFolderUrl = string.Join("/", folderNames, 1, folderNames.Length - 1);
                return ListEnsureFolder(folder, subFolderUrl);
            }

            return folder;
        }

        /// <summary>
        /// Upload a file to the specific library/folder
        /// </summary>
        /// <param name="onlineLibrary"></param>
        /// <param name="onlineLibraryFolder"></param>
        /// <param name="onlineFileName"></param>
        /// <returns></returns>
        public static string UploadFile(this List onlineLibrary, string onlineLibraryFolder, string onlineFileName)
        {
            var relativeUrl = string.Empty;
            var fileName = System.IO.Path.GetFileName(onlineFileName);
            try
            {
                var webUri = new Uri(onlineLibrary.Context.Url);

                var currentFolder = GetOrCreateFolder(onlineLibrary, onlineLibrary.RootFolder, onlineLibraryFolder);
                var logFileItem = GetFileInFolder(onlineLibrary, currentFolder, fileName);
                if (logFileItem == null)
                {
                    onlineLibrary.Context.Load(currentFolder, pf => pf.Name, pf => pf.Files, pf => pf.ServerRelativeUrl, pf => pf.TimeCreated);
                    onlineLibrary.Context.ExecuteQuery();

                    using (var stream = new System.IO.FileStream(onlineFileName, System.IO.FileMode.Open))
                    {

                        var creationInfo = new Microsoft.SharePoint.Client.FileCreationInformation();
                        creationInfo.Overwrite = true;
                        creationInfo.ContentStream = stream;
                        creationInfo.Url = fileName;

                        var uploadStatus = currentFolder.Files.Add(creationInfo);
                        onlineLibrary.Context.Load(uploadStatus, ups => ups.ServerRelativeUrl);
                        onlineLibrary.Context.ExecuteQuery();
                        relativeUrl = uploadStatus.ServerRelativeUrl;
                    }
                }
                else
                {
                    onlineLibrary.Context.Load(logFileItem.File, ups => ups.ServerRelativeUrl);
                    onlineLibrary.Context.ExecuteQuery();
                    relativeUrl = logFileItem.File.ServerRelativeUrl;
                }

                var assertedUri = new Uri(webUri, relativeUrl);
                relativeUrl = assertedUri.AbsoluteUri;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.TraceError("Failed to upload file {0} MSG:{1}", onlineFileName, ex.Message);
            }
            return relativeUrl;
        }

        /// <summary>
        /// Retreive file from the specified folder
        /// </summary>
        /// <param name="onlineLibrary"></param>
        /// <param name="destinationFolder"></param>
        /// <param name="onlineFileName"></param>
        /// <returns></returns>
        public static ListItem GetFileInFolder(this List onlineLibrary, Folder destinationFolder, string onlineFileName)
        {
            destinationFolder.Context.Load(destinationFolder, afold => afold.ServerRelativeUrl);
            destinationFolder.Context.ExecuteQuery();
            var relativeUrl = destinationFolder.ServerRelativeUrl;
            var context = destinationFolder.Context;
            try
            {
                CamlQuery camlQuery = new CamlQuery();
                var camlAndValue = CAML.And(
                            CAML.Eq(CAML.FieldValue("LinkFilename", FieldType.Text.ToString("f"), onlineFileName)),
                             CAML.Eq(CAML.FieldValue("FileDirRef", FieldType.Text.ToString("f"), relativeUrl)));

                camlQuery.ViewXml = CAML.ViewQuery(ViewScope.RecursiveAll,
                    CAML.Where(camlAndValue),
                    string.Empty,
                    CAML.ViewFields(CAML.FieldRef("Title")),
                    5);
                ListItemCollection listItems = onlineLibrary.GetItems(camlQuery);
                context.Load(listItems);
                context.ExecuteQuery();

                if (listItems.Count() > 0)
                {
                    var newItem = listItems.FirstOrDefault();
                    return newItem;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.TraceError("Failed to retrieve file {0} MSG:{1}", onlineFileName, ex.Message);
            }
            return null;
        }

        /// <summary>
        /// Query the list to retreive the last ID
        /// </summary>
        /// <param name="onlineLibrary">The List we will query</param>
        /// <param name="lastItemModifiedDate">The date of the last modified list item</param>
        /// <returns></returns>
        public static int QueryLastItemId(this List onlineLibrary, Nullable<DateTime> lastItemModifiedDate = null)
        {
            var returnId = 0;
            var camlFieldRefs = new string[] { "ID", "Created", "Modified", };
            var camlViewClause = CAML.ViewFields(camlFieldRefs.Select(s => CAML.FieldRef(s)).ToArray());
            var camlQuery = new CamlQuery()
            {
                ViewXml = CAML.ViewQuery(string.Empty, CAML.OrderBy(new OrderByField("Modified", false)), 10)
            };


            if (!lastItemModifiedDate.HasValue)
            {
                onlineLibrary.EnsureProperties(olp => olp.LastItemModifiedDate, olp => olp.LastItemUserModifiedDate);
                lastItemModifiedDate = onlineLibrary.LastItemModifiedDate;
            }


            ListItemCollectionPosition ListItemCollectionPosition = null;
            while (true)
            {
                camlQuery.ListItemCollectionPosition = ListItemCollectionPosition;
                var spListItems = onlineLibrary.GetItems(camlQuery);
                onlineLibrary.Context.Load(spListItems, lti => lti.ListItemCollectionPosition);
                onlineLibrary.Context.ExecuteQueryRetry();
                ListItemCollectionPosition = spListItems.ListItemCollectionPosition;

                if (spListItems.Any())
                {
                    foreach (var item in spListItems)
                    {
                        var itemModified = item.RetrieveListItemValue("Modified").ToDateTime();
                        System.Diagnostics.Trace.TraceInformation("Item {0} Modified {1} IS MATCH:{2}", item.Id, itemModified, (itemModified == lastItemModifiedDate));
                    }
                    returnId = spListItems.OrderByDescending(ob => ob.Id).FirstOrDefault().Id;
                    break;
                }

                if (ListItemCollectionPosition == null)
                {
                    break;
                }
            }

            return returnId;
        }

        /// <summary>
        /// Will evaluate the Library/List and if the list has reached the threshold it will produce SAFE queries
        /// </summary>
        /// <param name="onlineLibrary">The list to query</param>
        /// <param name="camlStatement">A base CAML query upon which the threshold query will be constructed</param>
        /// <returns>A collection of CAML queries NOT including WHERE</returns>
        public static List<string> SafeCamlClauseFromThreshold(this List onlineLibrary, string camlStatement = null, int? defaultLastItemId = default(int?))
        {
            var camlQueries = new List<string>();

            onlineLibrary.EnsureProperties(olp => olp.ItemCount, olp => olp.LastItemModifiedDate, olp => olp.LastItemUserModifiedDate);

            // we have reached a threshold and need to parse based on other criteria
            var itemCount = onlineLibrary.ItemCount;
            if (itemCount > 5000)
            {
                var lastItemId = (defaultLastItemId.HasValue) ? defaultLastItemId.Value : onlineLibrary.QueryLastItemId(onlineLibrary.LastItemModifiedDate);
                var startIdx = 0;
                var incrementor = 1000;

                for (var idx = startIdx; idx < lastItemId + 1;)
                {
                    var startsWithId = idx + 1;
                    var endsWithId = (idx + incrementor);
                    if (endsWithId >= lastItemId)
                    {
                        endsWithId = lastItemId + 1;
                    }

                    var thresholdEq = new SPThresholdEnumerationModel()
                    {
                        StartsWithId = startsWithId,
                        EndsWithId = endsWithId
                    };

                    var camlThresholdClause = thresholdEq.AndClause;
                    if (!string.IsNullOrEmpty(camlStatement))
                    {
                        camlThresholdClause = CAML.And(camlThresholdClause, camlStatement);
                    }
                    camlQueries.Add(camlThresholdClause);

                    idx += incrementor;
                }
            }
            else
            {
                camlQueries.Add(camlStatement ?? string.Empty);
            }

            return camlQueries;
        }

        /// <summary>
        /// get the xml for an xslt web part
        /// </summary>
        /// <param name="viewList">ID of the list</param>
        /// <param name="pageUrl">relative page url</param>
        /// <param name="title">title of the list</param>
        /// <param name="viewID">Represents the View base for the webpart</param>
        /// <returns>string</returns>
        public static string GetXsltWebPartXML(this List viewList, string pageUrl, string title, Guid viewID)
        {
            var executor = false;
            if (!viewList.IsPropertyAvailable("Id"))
            {
                viewList.Context.Load(viewList, vl => vl.Id, vl => vl.Title);
                executor = true;
            }
            if (!viewList.IsPropertyAvailable("RootFolder"))
            {
                viewList.Context.Load(viewList.RootFolder, rf => rf.ServerRelativeUrl, rf => rf.ItemCount, rf => rf.Name);
                executor = true;
            }

            // The properties were not loaded from caller
            if (executor)
            {
                viewList.Context.ExecuteQueryRetry();
            }

            Guid listID = viewList.Id;
            var listUrl = viewList.RootFolder.ServerRelativeUrl;


            StringBuilder wp = new StringBuilder(100);
            wp.Append("<?xml version=\"1.0\" encoding=\"utf-8\" ?>");
            wp.Append("<webParts>");
            wp.Append("	<webPart xmlns=\"http://schemas.microsoft.com/WebPart/v3\">");
            wp.Append("		<metaData>");
            wp.Append("			<type name=\"Microsoft.SharePoint.WebPartPages.XsltListViewWebPart, Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c\" />");
            wp.Append("			<importErrorMessage>Cannot import this Web Part.</importErrorMessage>");
            wp.Append("		</metaData>");
            wp.Append("		<data>");
            wp.Append("			<properties>");
            wp.Append("				<property name=\"Default\" type=\"bool\">False</property>");
            wp.Append("				<property name=\"IsIncluded\" type=\"bool\">True</property>");
            wp.Append("				<property name=\"NoDefaultStyle\" type=\"bool\">True</property>");
            wp.Append("				<property name=\"ViewContentTypeId\" type=\"string\" />");
            wp.AppendFormat("		<property name=\"ListUrl\" type=\"string\">{0}</property>", pageUrl);
            wp.AppendFormat("		<property name=\"ListId\" type=\"System.Guid, mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089\">{0}</property>", listID.ToString());
            wp.AppendFormat("		<property name=\"TitleUrl\" type=\"string\">{0}</property>", listUrl);
            wp.AppendFormat("		<property name=\"ListName\" type=\"string\">{0}</property>", listID.ToString("B").ToUpper());
            wp.AppendFormat("		<property name=\"Title\" type=\"string\">{0}</property>", title);
            wp.Append("             <property name=\"Toolbar Type\" type=\"string\">None</property>");
            wp.Append("				<property name=\"PageType\" type=\"Microsoft.SharePoint.PAGETYPE, Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c\">PAGE_NORMALVIEW</property>");
            wp.AppendFormat("       <property name=\"ViewGuid\" type=\"string\">{0}</property>", viewID.ToString("B").ToUpper());
            wp.Append("				<property name=\"XmlDefinition\" type=\"string\">");
            wp.AppendFormat("&lt;View Name=\"{1}\" Type=\"HTML\" Hidden=\"TRUE\" ReadOnly=\"TRUE\" OrderedView=\"TRUE\" DisplayName=\"\" Url=\"{0}\" Level=\"1\" BaseViewID=\"1\" ContentTypeID=\"0x\" &gt;&lt;Query&gt;&lt;OrderBy&gt;&lt;FieldRef Name=\"Title\" Ascending=\"TRUE\"/&gt;&lt;FieldRef Name=\"Modified\" Ascending=\"FALSE\"/&gt;&lt;/OrderBy&gt;&lt;/Query&gt;&lt;ViewFields&gt;&lt;FieldRef Name=\"Title\"/&gt;&lt;/ViewFields&gt;&lt;RowLimit Paged=\"TRUE\"&gt;30&lt;/RowLimit&gt;&lt;JSLink&gt;sp.ui.tileview.js&lt;/JSLink&gt;&lt;XslLink Default=\"TRUE\"&gt;main.xsl&lt;/XslLink&gt;&lt;Toolbar Type=\"None\"/&gt;&lt;/View&gt;",
                pageUrl, Guid.NewGuid().ToString("B").ToUpper());
            wp.Append("             </property>");
            wp.Append("			</properties>");
            wp.Append("		</data>");
            wp.Append("	</webPart>");
            wp.Append("</webParts>");
            return wp.ToString();
        }

        /// <summary>
        /// Adds or Updates an existing Custom Action [Url] into the [List] Custom Actions
        /// </summary>
        /// <param name="list"></param>
        /// <param name="customactionname"></param>
        /// <param name="commandUIExtension"></param>
        public static void AddOrUpdateCustomActionLink(this List list, string customactionname, string commandUIExtension, string location, int sequence)
        {
            var sitecustomActions = list.UserCustomActions;
            list.Context.Load(sitecustomActions);
            list.Context.ExecuteQueryRetry();

            UserCustomAction cssAction = null;
            if (sitecustomActions.Any(sa => sa.Name == customactionname))
            {
                cssAction = sitecustomActions.FirstOrDefault(fod => fod.Name == customactionname);
            }
            else
            {
                // Build a custom action
                cssAction = sitecustomActions.Add();
                cssAction.Name = customactionname;
            }

            cssAction.Sequence = sequence;
            cssAction.Location = location;
            cssAction.CommandUIExtension = commandUIExtension;
            cssAction.Update();
            list.Context.ExecuteQueryRetry();
        }

        /// <summary>
        /// Adds or Updates an existing Custom Action [Url] into the [List] Custom Actions
        /// </summary>
        /// <param name="list"></param>
        /// <param name="customactionname"></param>
        /// <param name="customactionurl"></param>
        /// <param name="title"></param>
        /// <param name="description"></param>
        /// <param name="location"></param>
        /// <param name="sequence">(default) 10000</param>
        /// <param name="groupName">(optional) adds custom group</param>
        public static void AddOrUpdateCustomActionLink(this List list, SPCustomActionList action)
        {
            var sitecustomActions = list.UserCustomActions;
            list.Context.Load(sitecustomActions);
            list.Context.ExecuteQueryRetry();

            UserCustomAction cssAction = null;
            if (sitecustomActions.Any(sa => sa.Name == action.name))
            {
                cssAction = sitecustomActions.FirstOrDefault(fod => fod.Name == action.name);
            }
            else
            {
                // Build a custom action
                cssAction = sitecustomActions.Add();
                cssAction.Name = action.name;
            }

            cssAction.Sequence = action.sequence;
            cssAction.Url = action.Url;
            cssAction.Description = action.Description;
            cssAction.Location = action.Location;
            cssAction.Title = action.Title;
            if (!string.IsNullOrEmpty(action.ImageUrl))
            {
                cssAction.ImageUrl = action.ImageUrl;
            }
            if (!string.IsNullOrEmpty(action.Group))
            {
                cssAction.Group = action.Group;
            }
            cssAction.Update();
            list.Context.ExecuteQueryRetry();
        }
    }
}
