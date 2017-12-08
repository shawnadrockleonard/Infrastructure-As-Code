using InfrastructureAsCode.Core.Models;
using InfrastructureAsCode.Core.Reports;
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
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace InfrastructureAsCode.Core.Extensions
{
    public static partial class ListExtensions
    {
        const string REGEX_INVALID_FILE_NAME_CHARS = @"[<>:;*?/\\|""&%\t\r\n]";

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
        public static Field CreateListColumn(this List hostList, SPFieldDefinitionModel fieldDefinition, ITraceLogger logger, List<SPGroupDefinitionModel> SiteGroups, List<SiteProvisionerFieldChoiceModel> provisionerChoices = null)
        {

            if (fieldDefinition == null)
            {
                throw new ArgumentNullException("fieldDef", "Field definition is required.");
            }

            if (fieldDefinition.LoadFromJSON && (provisionerChoices == null || !provisionerChoices.Any(pc => pc.FieldInternalName == fieldDefinition.InternalName)))
            {
                throw new ArgumentNullException("provisionerChoices", string.Format("You must specify a collection of field choices for the field {0}", fieldDefinition.Title));
            }

            // load fields into memory
            hostList.Context.Load(hostList.Fields, fc => fc.Include(f => f.Id, f => f.InternalName, fctx => fctx.Title, f => f.Title, f => f.JSLink, f => f.Indexed, f => f.CanBeDeleted, f => f.Required));
            hostList.Context.ExecuteQueryRetry();

            var returnField = hostList.Fields.FirstOrDefault(f => f.Id == fieldDefinition.FieldGuid || f.InternalName == fieldDefinition.InternalName);
            if (returnField == null)
            {
                try
                {
                    var baseFieldXml = hostList.CreateFieldDefinition(fieldDefinition, SiteGroups, provisionerChoices);
                    logger.LogInformation("Provision List {0} field {1} with XML:{2}", hostList.Title, fieldDefinition.InternalName, baseFieldXml);

                    // Should throw an exception if the field ID or Name exist in the list
                    var baseField = hostList.CreateField(baseFieldXml, fieldDefinition.AddToDefaultView, executeQuery: false);
                    hostList.Context.ExecuteQueryRetry();
                }
                catch (Exception ex)
                {
                    var msg = ex.Message;
                    logger.LogError(ex, "EXCEPTION: field {0} with message {1}", fieldDefinition.InternalName, msg);
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
        /// <param name="onlineLibrary">The target list for the folder</param>
        /// <param name="destinationFolder">The parent folder of the folter to be created</param>
        /// <param name="folderName">The folder to be created</param>
        /// <param name="defaultStartItemId">(OPTIONAL) If the list/library is above the threshold this will be the start index of the caml queries</param>
        /// <param name="defaultLastItemId">(OPTIONAL) If the list/library is above the threshold this will be the terminating index of the caml queries</param>
        /// <returns>The listitem as a folder</returns>
        public static Folder GetOrCreateFolder(this List onlineLibrary, Folder destinationFolder, string folderName, Nullable<int> defaultStartItemId = default(int?), Nullable<int> defaultLastItemId = default(int?))
        {
            if (!onlineLibrary.IsPropertyAvailable(lctx => lctx.BaseType))
            {
                onlineLibrary.Context.Load(onlineLibrary, lctx => lctx.BaseType, lctx => lctx.BaseTemplate);
                onlineLibrary.Context.ExecuteQueryRetry();
            }

            if (!destinationFolder.IsPropertyAvailable(fctx => fctx.ServerRelativeUrl)
                || !destinationFolder.IsObjectPropertyInstantiated(fctx => fctx.Folders))
            {
                destinationFolder.Context.Load(destinationFolder, afold => afold.ServerRelativeUrl, afold => afold.Folders);
                destinationFolder.Context.ExecuteQueryRetry();
            }


            ListItem folderItem = null;

            var folderRelativeUrl = destinationFolder.ServerRelativeUrl;
            var camlFields = new string[] { "Title", "ID" };
            var camlViewFields = CAML.ViewFields(camlFields.Select(s => CAML.FieldRef(s)).ToArray());

            // Remove invalid characters
            var trimmedFolder = HelperExtensions.GetCleanDirectory(folderName, string.Empty);
            var linkFileFilter = CAML.Eq(CAML.FieldValue("Title", FieldType.Text.ToString("f"), trimmedFolder));
            if (onlineLibrary.BaseType == BaseType.DocumentLibrary)
            {
                linkFileFilter = CAML.Or(
                    linkFileFilter,
                     CAML.Eq(CAML.FieldValue("LinkFilename", FieldType.Text.ToString("f"), trimmedFolder)));
            }

            var camlClause = CAML.And(
                CAML.Eq(CAML.FieldValue("FileDirRef", FieldType.Text.ToString("f"), folderRelativeUrl)),
                CAML.And(
                    CAML.Eq(CAML.FieldValue("FSObjType", FieldType.Integer.ToString("f"), 1.ToString())),
                    linkFileFilter
                )
            );

            var camlQueries = onlineLibrary.SafeCamlClauseFromThreshold(1000, camlClause, defaultStartItemId, defaultLastItemId);
            foreach (var camlAndValue in camlQueries)
            {
                ListItemCollectionPosition itemPosition = null;
                var camlWhereClause = CAML.Where(camlAndValue);
                var camlQuery = new CamlQuery()
                {
                    ViewXml = CAML.ViewQuery(
                        ViewScope.RecursiveAll,
                        camlWhereClause,
                        string.Empty,
                        camlViewFields,
                        5)
                };
                camlQuery.ListItemCollectionPosition = itemPosition;
                var listItems = onlineLibrary.GetItems(camlQuery);
                onlineLibrary.Context.Load(listItems, lti => lti.ListItemCollectionPosition);
                onlineLibrary.Context.ExecuteQueryRetry();

                if (listItems.Count() > 0)
                {
                    folderItem = listItems.FirstOrDefault();
                    System.Diagnostics.Trace.TraceInformation("Folder {0} exists in the destination folder.  Skip folder creation .....", folderName);
                    break;
                }
            };

            if (folderItem != null)
            {
                return folderItem.Folder;
            }

            try
            {
                var info = new ListItemCreationInformation
                {
                    UnderlyingObjectType = FileSystemObjectType.Folder,
                    LeafName = trimmedFolder,
                    FolderUrl = folderRelativeUrl
                };

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
            if (!parentList.IsObjectPropertyInstantiated(pctx => pctx.RootFolder))
            {
                parentList.Context.Load(parentList, pl => pl.RootFolder, pl => pl.RootFolder.ServerRelativeUrl);
                parentList.Context.ExecuteQueryRetry();
            }

            var listUri = new Uri(parentList.RootFolder.ServerRelativeUrl);
            var relativeUri = listUri.MakeRelativeUri(new Uri(folderUrl));
            var relativeUrl = folderUrl.Replace(parentList.RootFolder.ServerRelativeUrl, "");

            var folder = ListEnsureFolder(parentList.RootFolder, folderUrl);
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
            if (!parentFolder.IsObjectPropertyInstantiated(pctx => pctx.Folders))
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
        /// Upload a file from disk to the specific <paramref name="onlineFolder"/>
        /// </summary>
        /// <param name="onlineLibrary"></param>
        /// <param name="onlineFolder"></param>
        /// <param name="fileNameWithPath"></param>
        /// <param name="clobber">(OPTIONAL) if true then overwrite the existing file</param>
        /// <exception cref="System.IO.FileNotFoundException">File not found if fullfilename does not exist</exception>
        /// <returns></returns>
        public static string UploadFile(this List onlineLibrary, Folder onlineFolder, string fileNameWithPath, bool clobber = false)
        {
            var relativeUrl = string.Empty;
            var fileName = System.IO.Path.GetFileName(fileNameWithPath);
            if (!System.IO.File.Exists(fileNameWithPath))
            {
                throw new System.IO.FileNotFoundException(string.Format("File {0} does not exists on disk", fileNameWithPath));
            }

            if (!onlineFolder.IsPropertyAvailable(fctx => fctx.ServerRelativeUrl))
            {
                try
                {
                    // setup processing of folder in the parent folder
                    onlineFolder.Context.Load(onlineFolder, pf => pf.ServerRelativeUrl);
                    onlineFolder.Context.ExecuteQueryRetry();
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Trace.TraceError("Failed to ensure folder server relative url property MSG:{0}", ex.Message);
                }
            }

            var logFileItem = GetFileInFolder(onlineLibrary, onlineFolder, fileName);
            if (logFileItem == null
                || (clobber && logFileItem != null))
            {
                onlineLibrary.Context.Load(onlineFolder, pf => pf.Name, pf => pf.Files, pf => pf.ServerRelativeUrl, pf => pf.TimeCreated);
                onlineLibrary.Context.ExecuteQueryRetry();

                using (var stream = new System.IO.FileStream(fileNameWithPath, System.IO.FileMode.Open))
                {

                    var creationInfo = new Microsoft.SharePoint.Client.FileCreationInformation
                    {
                        Overwrite = clobber,
                        ContentStream = stream,
                        Url = fileName
                    };

                    var uploadStatus = onlineFolder.Files.Add(creationInfo);
                    onlineLibrary.Context.Load(uploadStatus, ups => ups.ServerRelativeUrl);
                    onlineLibrary.Context.ExecuteQueryRetry();
                    relativeUrl = uploadStatus.ServerRelativeUrl;
                }
            }
            else
            {
                onlineLibrary.Context.Load(logFileItem.File, ups => ups.ServerRelativeUrl);
                onlineLibrary.Context.ExecuteQueryRetry();
                relativeUrl = logFileItem.File.ServerRelativeUrl;
            }

            var webUri = new Uri(onlineFolder.Context.Url);
            var assertedUri = new Uri(webUri, relativeUrl);
            relativeUrl = assertedUri.AbsoluteUri;

            return relativeUrl;
        }

        /// <summary>
        /// Upload a file from disk to the specific <paramref name="onlineLibraryFolder"/> inside the RootFolder
        /// </summary>
        /// <param name="onlineLibrary"></param>
        /// <param name="onlineLibraryFolder"></param>
        /// <param name="fileNameWithPath"></param>
        /// <param name="clobber">(OPTIONAL) if true then overwrite the existing file</param>
        /// <exception cref="System.IO.FileNotFoundException">File not found if fullfilename does not exist</exception>
        /// <returns></returns>
        public static string UploadFile(this List onlineLibrary, string onlineLibraryFolder, string fileNameWithPath, bool clobber = false)
        {
            var fileName = System.IO.Path.GetFileName(fileNameWithPath);
            if (!System.IO.File.Exists(fileNameWithPath))
            {
                throw new System.IO.FileNotFoundException(string.Format("File {0} does not exists on disk", fileNameWithPath));
            }

            if (!onlineLibrary.IsObjectPropertyInstantiated(fctx => fctx.RootFolder)
                && !onlineLibrary.RootFolder.IsPropertyAvailable(fctx => fctx.ServerRelativeUrl))
            {
                try
                {
                    // setup processing of folder in the parent folder
                    onlineLibrary.Context.Load(onlineLibrary, pf => pf.RootFolder, pf => pf.RootFolder.ServerRelativeUrl);
                    onlineLibrary.Context.ExecuteQueryRetry();
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Trace.TraceError("Failed to ensure root folder propert MSG:{0}", ex.Message);
                }
            }

            var currentFolder = GetOrCreateFolder(onlineLibrary, onlineLibrary.RootFolder, onlineLibraryFolder);
            var relativeUrl = onlineLibrary.UploadFile(currentFolder, fileNameWithPath, clobber);
            return relativeUrl;
        }

        /// <summary>
        /// Upload a File via the REST API interface
        /// </summary>
        /// <param name="context">The context established by AppPrincipal</param>
        /// <param name="relativeUrl">The folder structure, relative to the Web</param>
        /// <param name="fileWithPath">The full file name with path on local disk</param>
        /// <param name="ensureFolder">(OPTIONAL) true will ensure the relativeUrl path exists in SharePoint</param>
        /// <returns></returns>
        /// <remarks>
        /// The <paramref name="context"/> should be an AppPrincipal context which will contain the bearer token for OAuth interactions pulled from the SharePointContext
        /// </remarks>
        public static bool UploadFileViaREST(this ClientContext context, string relativeUrl, string fileWithPath, bool ensureFolder = false)
        {
            if (!System.IO.File.Exists(fileWithPath))
            {
                throw new ArgumentException(string.Format("The file {0} does not exist on disc.", fileWithPath));
            }

            if (context.Web.RootFolder.ServerObjectIsNull())
            {
                context.Load(context.Web, ctx => ctx.ServerRelativeUrl, ctx => ctx.RootFolder, ctx => ctx.RootFolder.ServerRelativeUrl);
                context.ExecuteQueryRetry();
            }

            if (ensureFolder)
            {
                var folderPath = context.Web.RootFolder.ListEnsureFolder(relativeUrl);
                if (folderPath == null || folderPath.ServerObjectIsNull())
                {
                    throw new Exception("Failed to ensure folder path directories.");
                }
            }

            var accessToken = string.Empty;

            try
            {
                accessToken = context.GetAccessToken();
            }
            catch (Exception ex)
            {
                throw new Exception("Failed to retreive Access Token", ex);
            }

            try
            {
                var fileName = System.IO.Path.GetFileName(fileWithPath);
                var fileBuffer = System.IO.File.ReadAllBytes(fileWithPath);
                var fileSize = fileBuffer.Length;

                var strURL = string.Format("{0}/_api/web/GetFolderByServerRelativeUrl('{1}')/Files/add(url='{2}',overwrite=true)",
                     context.Url,
                     relativeUrl,
                     fileName);

                var wreq = System.Net.HttpWebRequest.Create(strURL) as System.Net.HttpWebRequest;
                wreq.UseDefaultCredentials = true;

                // Upload to SharePoiint
                var authToken = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);
                wreq.Headers.Add(System.Net.HttpRequestHeader.Authorization, authToken.ToString());
                wreq.Method = "POST";
                wreq.Timeout = 1000000;
                wreq.Accept = "application/json; odata=verbose";
                wreq.ContentLength = fileSize;

                using (var sRequest = wreq.GetRequestStream())
                {
                    sRequest.Write(fileBuffer, 0, fileSize);
                }

                using (var wresp = wreq.GetResponse())
                {
                    var response = (System.Net.HttpWebResponse)wresp;


                    using (var sr = new System.IO.StreamReader(wresp.GetResponseStream()))
                    {
                        //var webmsg = sr.ReadToEnd();
                        System.Diagnostics.Trace.TraceInformation("Server response {0} with description {1}", response.StatusCode, response.StatusDescription);

                    }
                }

                return true;
            }
            catch (System.Net.WebException e)
            {
                if (e.Status == System.Net.WebExceptionStatus.ProtocolError)
                {
                    var response = (System.Net.HttpWebResponse)e.Response;
                    throw new Exception(string.Format("Errorcode: {0}", (int)response.StatusCode), e);
                }
                else
                {
                    throw new Exception(string.Format("Error: {0}", e.Status), e);
                }
            }
            catch (Exception exError)
            {
                //Log Error // Catch Folder Creation exceptions
                throw (exError);
            }
        }

        /// <summary>
        /// Retreive file from the specified folder
        /// </summary>
        /// <param name="onlineLibrary">The library hosting the file</param>
        /// <param name="destinationFolder">The folder where the file is located</param>
        /// <param name="onlineFileName">The filename</param>
        /// <returns></returns>
        public static ListItem GetFileInFolder(this List onlineLibrary, Folder destinationFolder, string onlineFileName)
        {
            if (!destinationFolder.IsPropertyAvailable(fctx => fctx.ServerRelativeUrl))
            {
                destinationFolder.Context.Load(destinationFolder, afold => afold.ServerRelativeUrl);
                destinationFolder.Context.ExecuteQueryRetry();
            }

            var relativeUrl = destinationFolder.ServerRelativeUrl;

            try
            {
                var camlAndValue = CAML.And(
                            CAML.Eq(CAML.FieldValue("LinkFilename", FieldType.Text.ToString("f"), onlineFileName)),
                             CAML.Eq(CAML.FieldValue("FileDirRef", FieldType.Text.ToString("f"), relativeUrl)));

                var camlQuery = new CamlQuery()
                {
                    ViewXml = CAML.ViewQuery(ViewScope.RecursiveAll,
                    CAML.Where(camlAndValue),
                    string.Empty,
                    CAML.ViewFields(CAML.FieldRef("Title")),
                    5)
                };

                var listItems = onlineLibrary.GetItems(camlQuery);
                destinationFolder.Context.Load(listItems);
                destinationFolder.Context.ExecuteQueryRetry();

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
        /// Moves a file from <paramref name="serverRelativeUrl"/> to <paramref name="newRelativeUrl"/>
        /// </summary>
        /// <param name="onlineLibrary">The library/list where the item resides</param>
        /// <param name="serverRelativeUrl">The current path for the list item</param>
        /// <param name="newRelativeUrl">The target path for the list item</param>
        /// <returns></returns>
        public static bool MoveFileToFolder(this List onlineLibrary, string serverRelativeUrl, string newRelativeUrl)
        {
            var context = onlineLibrary.Context;
            try
            {
                var targetItem = onlineLibrary.ParentWeb.GetFileByServerRelativeUrl(serverRelativeUrl);
                context.Load(targetItem);
                context.ExecuteQueryRetry();

                targetItem.MoveTo(newRelativeUrl, MoveOperations.None);
                context.ExecuteQueryRetry();
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.TraceError(ex.Message);
            }
            return false;
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
                ViewXml = CAML.ViewQuery(
                    ViewScope.RecursiveAll,
                    string.Empty,
                    CAML.OrderBy(new OrderByField("ID", false)),
                    camlViewClause,
                    1)
            };


            ListItemCollectionPosition itemPosition = null;
            while (true)
            {
                camlQuery.ListItemCollectionPosition = itemPosition;
                var spListItems = onlineLibrary.GetItems(camlQuery);
                onlineLibrary.Context.Load(spListItems, lti => lti.ListItemCollectionPosition);
                onlineLibrary.Context.ExecuteQueryRetry();
                itemPosition = spListItems.ListItemCollectionPosition;

                if (spListItems.Any())
                {
                    if (lastItemModifiedDate.HasValue)
                    {
                        foreach (var item in spListItems)
                        {
                            var itemModified = item.RetrieveListItemValue("Modified").ToNullableDatetime();
                            if (itemModified == lastItemModifiedDate)
                            {
                                returnId = item.Id;
                                break;
                            }
                        }
                    }
                    else
                    {
                        returnId = spListItems.OrderByDescending(ob => ob.Id).FirstOrDefault().Id;
                    }

                    if (returnId > 0)
                    {
                        // Found the item ID that matches the specified date - return it
                        break;
                    }
                }

                if (itemPosition == null)
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
        /// <param name="incrementor">(Default) 1000 rows in the query, can specify up to 5000</param>
        /// <param name="camlStatement">A base CAML query upon which the threshold query will be constructed</param>
        /// <param name="defaultStartItemId">(OPTIONAL) if specified the caml queries will begin at ID >= this value</param>
        /// <param name="defaultLastItemId">(OPTIONAL) if specified the caml queries will terminate at this value; if not specified a query will be executed to retreive the lastitemid</param>
        /// <returns>A collection of CAML queries NOT including WHERE</returns>
        public static List<string> SafeCamlClauseFromThreshold(this List onlineLibrary, int incrementor = 1000, string camlStatement = null, Nullable<int> defaultStartItemId = default(Nullable<int>), Nullable<int> defaultLastItemId = default(Nullable<int>))
        {
            if (incrementor > 5000)
            {
                throw new InvalidOperationException(string.Format("CAML Queries must return fewer than 5000 rows, you specified {0}", incrementor));
            }

            var camlQueries = new List<string>();
            var lastItemId = 0;
            var startIdx = 0;

            // Check if the List/Library ItemCount exists
            if (!onlineLibrary.IsPropertyAvailable(octx => octx.ItemCount))
            {
                onlineLibrary.Context.Load(onlineLibrary, octx => octx.ItemCount);
                onlineLibrary.Context.ExecuteQueryRetry();
            }

            // we have reached a threshold and need to parse based on other criteria
            var itemCount = onlineLibrary.ItemCount;
            if (itemCount > 5000)
            {
                if (defaultStartItemId.HasValue)
                {
                    startIdx = defaultStartItemId.Value;
                }

                if (defaultLastItemId.HasValue)
                {
                    lastItemId = defaultLastItemId.Value;
                }
                else
                {

                    lastItemId = onlineLibrary.QueryLastItemId();
                }

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
            if (!viewList.IsPropertyAvailable(lctx => lctx.Id))
            {
                viewList.Context.Load(viewList, vl => vl.Id, vl => vl.Title);
                executor = true;
            }
            if (!viewList.IsPropertyAvailable(lctx => lctx.RootFolder))
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

        /// <summary>
        /// Parses the string value into an <see cref="Microsoft.SharePoint.Client.ViewType"/>
        /// </summary>
        /// <param name="viewType"></param>
        /// <returns></returns>
        public static ViewType TryGetViewType(string viewType)
        {
            ViewType viewCamlType = ViewType.None;
            if (!string.IsNullOrEmpty(viewType))
            {
                foreach (var vtype in Enum.GetNames(typeof(ViewType)))
                {
                    if (vtype.Equals(viewType, StringComparison.InvariantCultureIgnoreCase))
                    {
                        viewCamlType = (ViewType)Enum.Parse(typeof(ViewType), vtype);
                        break;
                    }
                }
            }
            return viewCamlType;
        }

        /// <summary>
        /// Get List Template Type
        /// </summary>
        /// <param name="list">List template CSOM</param>
        /// <returns>returns List template type </returns>
        public static ListTemplateType GetListTemplateType(this List list)
        {
            try
            {
                return (ListTemplateType)Enum.Parse(typeof(ListTemplateType), list.BaseTemplate.ToString());
            }
            catch
            {
                throw new System.ComponentModel.InvalidEnumArgumentException("ListTemplateType", list.BaseTemplate, typeof(ListTemplateType));
            }
        }
    }
}
