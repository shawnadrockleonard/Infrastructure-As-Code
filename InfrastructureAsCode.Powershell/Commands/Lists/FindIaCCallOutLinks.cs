using InfrastructureAsCode.Core.Extensions;
using InfrastructureAsCode.Powershell.CmdLets;
using InfrastructureAsCode.Powershell.Extensions;
using InfrastructureAsCode.Powershell.Models;
using InfrastructureAsCode.Powershell.PipeBinds;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Powershell.Lists
{
    /// <summary>
    /// CmdLet will query the list for CSOM properties
    /// </summary>
    /// <remarks>
    ///     Find-IaCCallOutLinks -List "Sample List" -PartialUrl "onpremhostheader" -Path "FolderDir" -Verbose  
    /// </remarks>
    [Cmdlet(VerbsCommon.Find, "IaCCallOutLinks", SupportsShouldProcess = false)]
    [OutputType(typeof(List<CalloutLinkModel>))]
    public class FindIaCCallOutLinks : IaCCmdlet
    {
        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
        public ListPipeBind List { get; set; }

        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 1)]
        public string PartialUrl { get; set; }

        /// <summary>
        /// View Identity
        /// </summary>
        [Parameter(Mandatory = false, ValueFromPipeline = true, Position = 2)]
        public string Path { get; set; }

        [Parameter(Mandatory = false, ValueFromPipeline = true, Position = 3)]
        public Nullable<int> EndId { get; set; }


        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();
            LogVerbose("Scanning CSOM callout");

            var invalidLinks = new List<CalloutLinkModel>();

            var fields = new string[]
            {
                "_dlc_DocId", "_dlc_DocIdUrl", "Modified", "Editor", "ServerRedirectedEmbedUri", "FileRef", "Title"
            };
            var fieldsXml = CAML.ViewFields(fields.Select(s => CAML.FieldRef(s)).ToArray());

            var onlineLibrary = List.GetList(this.ClientContext.Web);

            if (!EndId.HasValue)
            {
                onlineLibrary.EnsureProperties(ol => ol.ItemCount);
                EndId = onlineLibrary.ItemCount;
            }
            ListItemCollectionPosition itemCollectionPosition = null;
            CamlQuery camlQuery = new CamlQuery();

            for (var idx = 1; idx <= EndId; idx += 1000)
            {
                camlQuery.ViewXml = string.Format(@"<View Scope='RecursiveAll'><Query>
    <OrderBy><FieldRef Name='ID' /></OrderBy>
    <Where>
        <And>
            <And>
                {0}{1}
            </And>
            <Or>
                {2}{3}
            </Or>
        </And>
    </Where>
    {4}
    <RowLimit Paged='TRUE'>30</RowLimit>
</Query></View>",
        CAML.Geq(CAML.FieldValue("ID", FieldType.Integer.ToString("f"), idx.ToString())),
        CAML.Leq(CAML.FieldValue("ID", FieldType.Integer.ToString("f"), (idx + 1000).ToString())),
        CAML.Eq(CAML.FieldValue("FileDirRef", FieldType.Text.ToString("f"), Path)),
        CAML.Contains(CAML.FieldValue("_dlc_DocIdUrl", FieldType.URL.ToString("f"), PartialUrl)),
        fieldsXml
        );

                while (true)
                {
                    camlQuery.ListItemCollectionPosition = itemCollectionPosition;
                    ListItemCollection listItems = onlineLibrary.GetItems(camlQuery);
                    this.ClientContext.Load(listItems);
                    this.ClientContext.ExecuteQuery();
                    itemCollectionPosition = listItems.ListItemCollectionPosition;
                    if (listItems.Count() > 0)
                    {
                        foreach (var listItem in listItems)
                        {
                            var item = onlineLibrary.GetItemById(listItem.Id);
                            this.ClientContext.Load(item);
                            this.ClientContext.ExecuteQuery();

                            var docId = item.RetrieveListItemValue("_dlc_DocId");
                            var docIdUrl = item.RetrieveListItemValueAsHyperlink("_dlc_DocIdUrl");
                            var modified = item.RetrieveListItemValue("Modified").ToDate();
                            var editor = item.RetrieveListItemUserValue("Editor");
                            var redirectEmbeddedUrl = item.RetrieveListItemValue("ServerRedirectedEmbedUri");
                            var fileRef = item.RetrieveListItemValue("FileRef");
                            var title = item.RetrieveListItemValue("Title");

                            invalidLinks.Add(new CalloutLinkModel()
                            {
                                DocId = docId,
                                DocIdUrl = (docIdUrl == null) ? string.Empty : docIdUrl.Url,
                                Modified = modified,
                                EditorEmail = (editor == null) ? string.Empty : editor.ToUserEmailValue(),
                                EmbeddedUrl = redirectEmbeddedUrl,
                                FileUrl = fileRef,
                                Title = title,
                                Id = listItem.Id
                            });
                        }
                    }

                    if (itemCollectionPosition == null)
                    {
                        break;
                    }
                }
            }

            WriteObject(invalidLinks);
        }
    }
}
