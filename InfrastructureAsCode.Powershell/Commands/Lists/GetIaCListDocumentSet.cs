using InfrastructureAsCode.Core;
using InfrastructureAsCode.Core.Constants;
using InfrastructureAsCode.Core.Extensions;
using InfrastructureAsCode.Powershell.CmdLets;
using InfrastructureAsCode.Powershell.PipeBinds;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.DocumentSet;
using Microsoft.SharePoint.Client.Utilities;
using OfficeDevPnP.Core.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Management.Automation;
using System.Text;
using System.Xml.Linq;

namespace InfrastructureAsCode.Powershell.Commands.Lists
{
    /// <summary>
    /// CmdLet will provide a sample to query a SPView and iterate over the results
    /// </summary>
    [Cmdlet(VerbsCommon.Get, "IaCListDocumentSet", SupportsShouldProcess = false)]
    public class GetIaCListDocumentSet : IaCCmdlet
    {
        /// <summary>
        /// View Identity
        /// </summary>
        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
        public ViewPipeBind List { get; set; }

        /// <summary>
        /// View Identity
        /// </summary>
        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 1)]
        public ViewPipeBind View { get; set; }

        /// <summary>
        /// Directory where files will be written
        /// </summary>
        [Parameter(Mandatory = true)]
        public string TargetLocation { get; set; }

        /// <summary>
        /// contains columns from view for inference
        /// </summary>
        private List<FieldMappings> ColumnMappings { get; set; }

        /// <summary>
        /// Pre-process to evaluate the directory location
        /// </summary>
        protected override void OnBeginInitialize()
        {
            if (!System.IO.Directory.Exists(TargetLocation))
            {
                throw new System.IO.DirectoryNotFoundException(string.Format("Directory {0} does not exists", TargetLocation));
            }
        }

        /// <summary>
        /// Processing
        /// </summary>
        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();


            var paramlistname = List.Title;
            var paramviewname = View.Title;
            this.ColumnMappings = new List<FieldMappings>();
            var documentsetcontenttypeid = "0x0120D520";


            try
            {
                var viewlist = ClientContext.Web.GetListByTitle(paramlistname);
                ClientContext.Load(viewlist, rcll => rcll.Fields, rcll => rcll.ItemCount, rcll => rcll.ContentTypes, rcll => rcll.BaseType, rcll => rcll.RootFolder);
                ClientContext.Load(viewlist.Views, wv => wv.Include(wvi => wvi.Title, wvi => wvi.Id, wvi => wvi.ListViewXml, wvi => wvi.ViewFields));
                ClientContext.ExecuteQueryRetry();



                var viewFieldHeaderHtml = string.Empty;
                var view = viewlist.Views.FirstOrDefault(w => w.Title.Trim().Equals(paramviewname, StringComparison.CurrentCultureIgnoreCase));
                if (view != null)
                {
                    var doc = XDocument.Parse(view.ListViewXml);

                    var queryXml = doc.Root.Element(XName.Get("Query"));
                    var camlFieldXml = doc.Root.Element(XName.Get("ViewFields"));
                    var queryWhereXml = queryXml.Element(XName.Get("Where"));
                    var queryGroupByXml = queryXml.Element(XName.Get("GroupBy"));
                    var queryOrderXml = queryXml.Element(XName.Get("OrderBy"));

                    var queryViewCaml = ((camlFieldXml != null) ? camlFieldXml.ToString() : string.Empty);
                    var queryWhereCaml = ((queryWhereXml != null) ? queryWhereXml.ToString() : string.Empty);
                    var queryOrderCaml = ((queryOrderXml != null) ? queryOrderXml.ToString() : string.Empty);
                    var viewFields = new List<string>() { "ContentTypeId", "FileRef", "FileDirRef", "FileLeafRef" };

                    if (viewlist.BaseType == BaseType.GenericList)
                    {
                        viewFields.AddRange(new string[] { ConstantsListFields.Field_LinkTitle, ConstantsListFields.Field_LinkTitleNoMenu });
                    }

                    if (viewlist.BaseType == BaseType.DocumentLibrary)
                    {
                        viewFields.AddRange(new string[] { ConstantsLibraryFields.Field_LinkFilename, ConstantsLibraryFields.Field_LinkFilenameNoMenu });
                    }


                    foreach (var xnode in camlFieldXml.Descendants())
                    {
                        var attributeValue = xnode.Attribute(XName.Get("Name"));
                        var fe = attributeValue.Value;
                        if (!viewFields.Any(vf => vf == fe))
                        {
                            viewFields.Add(fe);
                        }
                    }

                    // lets override the view field XML with some additional columns
                    queryViewCaml = CAML.ViewFields(viewFields.Select(s => CAML.FieldRef(s)).ToArray());


                    view.ViewFields.ToList().ForEach(fe =>
                    {
                        var fieldDisplayName = viewlist.Fields.FirstOrDefault(fod => fod.InternalName == fe);

                        ColumnMappings.Add(new FieldMappings()
                        {
                            columnInternalName = fieldDisplayName.InternalName,
                            columnMandatory = fieldDisplayName.Required,
                            columnType = fieldDisplayName.FieldTypeKind
                        });
                    });


                    var camlQueryXml = CAML.ViewQuery(ViewScope.RecursiveAll, queryWhereCaml, queryOrderCaml, queryViewCaml, 500);

                    ListItemCollectionPosition camlListItemCollectionPosition = null;
                    var camlQuery = new CamlQuery();
                    camlQuery.ViewXml = camlQueryXml;


                    while (true)
                    {
                        camlQuery.ListItemCollectionPosition = camlListItemCollectionPosition;
                        var spListItems = viewlist.GetItems(camlQuery);
                        this.ClientContext.Load(spListItems, lti => lti.ListItemCollectionPosition);
                        this.ClientContext.ExecuteQueryRetry();
                        camlListItemCollectionPosition = spListItems.ListItemCollectionPosition;

                        foreach (var spItem in spListItems)
                        {
                            var fileurl = spItem.RetrieveListItemValue(ConstantsFields.Field_FileRef);
                            var progid = spItem.RetrieveListItemValue("ProgId");

                            LogVerbose("Item {0} ProgId {1} URL:{2}", spItem.Id, progid, fileurl);

                            var contenttypeid = spItem.RetrieveListItemValue("ContentTypeId");
                            if (contenttypeid.StartsWith(documentsetcontenttypeid) || progid.Equals("Sharepoint.DocumentSet", StringComparison.CurrentCultureIgnoreCase))
                            {
                                // process the docset
                                CheckDocumentSetMapping(viewlist, fileurl);

                                // process items inside the docset
                                CheckDocumentsByCaml(viewlist, viewFields, fileurl);
                            }
                        }

                        if (camlListItemCollectionPosition == null)
                        {
                            break;
                        }
                    }
                }

            }
            catch (Exception fex)
            {
                LogError(fex, "Failed to parse view and produce HTML report");
            }


        }

        internal void CheckDocumentSetMapping(List docsetlist, string relativeUrl)
        {

            Folder docsetfolder = ClientContext.Web.GetFolderByServerRelativeUrl(relativeUrl);
            ClientContext.Load(docsetfolder, fld => fld.Name, fld => fld.ParentFolder.Name, fld => fld.ServerRelativeUrl, fld => fld.ListItemAllFields);
            ClientContext.ExecuteQueryRetry();

            var onlineurl = relativeUrl.Replace(docsetlist.RootFolder.ServerRelativeUrl, "");
            var localdocuset = FullDocumentSetPath(this.TargetLocation, onlineurl, docsetfolder.Name);

            DocumentSet docSet = DocumentSet.GetDocumentSet(ClientContext, docsetfolder);
            var docSetStream = docSet.ExportDocumentSet();
            ClientContext.ExecuteQueryRetry();
            using (var fs = new System.IO.FileStream(localdocuset.FullName, System.IO.FileMode.Create, System.IO.FileAccess.Write, System.IO.FileShare.None))
            {
                docSetStream.Value.CopyTo(fs);
            }
        }

        internal void CheckDocumentsByCaml(List docsetlist, List<string> docsetfields, string relativeUrl)
        {
            var onlineurl = relativeUrl.Replace(docsetlist.RootFolder.ServerRelativeUrl, "");

            var query = Microsoft.SharePoint.Client.CamlQuery.CreateAllItemsQuery(100, docsetfields.ToArray());
            query.FolderServerRelativeUrl = relativeUrl;
            var allItems = docsetlist.GetItems(query);
            docsetlist.Context.Load(allItems);
            docsetlist.Context.ExecuteQueryRetry();

            foreach (var spItem in allItems)
            {
                var fileurl = spItem.RetrieveListItemValue(ConstantsFields.Field_FileRef);
                var filedownloaded = spItem.RetrieveListItemValue("Downloaded").ToBoolean();
                var fileinfo = new System.IO.FileInfo(fileurl);
                var filewithpath = FullDocumentPath(this.TargetLocation, onlineurl, fileinfo.Name);
                LogVerbose("Item {0} Downloaded {1} URL {2}", spItem.Id, filedownloaded, fileurl);

                using (var openFile = Microsoft.SharePoint.Client.File.OpenBinaryDirect(ClientContext, fileurl.ToString()))
                {
                    using (var fileStream = new System.IO.FileStream(filewithpath.FullName, System.IO.FileMode.Create))
                    {
                        openFile.Stream.CopyTo(fileStream);
                        fileStream.Close();
                    }
                }
            }
        }

        internal System.IO.FileInfo FullDocumentSetPath(string targetLocation, string partialUrl, string folderName)
        {
            var local = new System.IO.DirectoryInfo(targetLocation);
            local = local.CreateSubdirectory("DocumentSets");
            foreach (var partial in partialUrl.Split(new string[] { "/" }, StringSplitOptions.RemoveEmptyEntries))
            {
                local = local.CreateSubdirectory(partial, local.GetAccessControl());
            }

            var localfile = new System.IO.FileInfo(string.Format("{0}\\{1}.zip", local.FullName, folderName));
            return localfile;
        }

        internal System.IO.FileInfo FullDocumentPath(string targetLocation, string partialUrl, string fileName)
        {
            var local = new System.IO.DirectoryInfo(targetLocation);
            local = local.CreateSubdirectory("FileStream");
            foreach (var partial in partialUrl.Split(new string[] { "/" }, StringSplitOptions.RemoveEmptyEntries))
            {
                local = local.CreateSubdirectory(partial, local.GetAccessControl());
            }

            var localfile = new System.IO.FileInfo(string.Format("{0}\\{1}", local.FullName, fileName));
            return localfile;
        }

        internal class FieldMappings
        {
            internal Microsoft.SharePoint.Client.FieldType columnType { get; set; }

            internal string columnInternalName { get; set; }

            internal bool columnMandatory { get; set; }
        }
    }
}
