using InfrastructureAsCode.Core.Extensions;
using InfrastructureAsCode.Powershell.Commands.Base;
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
    /// CmdLet will query the list and update the docid properties
    /// </summary>
    /// <remarks>
    ///     Set-IaCCallOutLinksByObjects -List "List name" -Items @({}) -PartialUrl "onpremhostheader" -Verbose  
    /// </remarks>
    [Cmdlet(VerbsCommon.Set, "IaCCallOutLinksByObjects", SupportsShouldProcess = false)]
    public class SetIaCCallOutLinksByObjects : IaCCmdlet
    {
        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
        public ListPipeBind List { get; set; }

        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 1)]
        public List<CalloutLinkModel> Items { get; set; }

        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 2)]
        public string PartialUrl { get; set; }


        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();


            var fields = new string[]
            {
                "_dlc_DocId", "_dlc_DocIdUrl", "Modified", "Editor", "ServerRedirectedEmbedUri", "FileRef", "Title"
            };
            var fieldsXml = CAML.ViewFields(fields.Select(s => CAML.FieldRef(s)).ToArray());

            var onlineLibrary = List.GetList(this.ClientContext.Web);

            foreach (var csvitem in Items)
            {
                var ItemId = csvitem.Id;
                try
                {
                    var item = onlineLibrary.GetItemById(ItemId);
                    this.ClientContext.Load(item);
                    this.ClientContext.ExecuteQueryRetry();

                    if (item == null)
                    {
                        LogWarning("Failed to retreive listitem ID {0}", ItemId);
                    }
                    else
                    {
                        var docId = item.RetrieveListItemValue("_dlc_DocId");
                        var docIdUrl = item.RetrieveListItemValueAsHyperlink("_dlc_DocIdUrl");
                        var modified = item.RetrieveListItemValue("Modified");
                        var editor = item.RetrieveListItemUserValue("Editor");
                        var redirectEmbeddedUrl = item.RetrieveListItemValue("ServerRedirectedEmbedUri");
                        var fileRef = item.RetrieveListItemValue("FileRef");
                        var title = item.RetrieveListItemValue("Title");
                        LogVerbose("[PRE UPDATE] ==> DocId {0}  DocIdUrl {1}, Modified {2} Editor {3}, Embedded Url {4}, FileRef {5}", docId, (docIdUrl != null ? docIdUrl.Url : ""), modified, editor.Email, redirectEmbeddedUrl, fileRef);

                        if (docIdUrl != null && docIdUrl.Url.IndexOf(PartialUrl, StringComparison.InvariantCultureIgnoreCase) > -1)
                        {

                            item["_dlc_DocId"] = null;
                            item["_dlc_DocIdUrl"] = null;
                            item["Modified"] = modified;
                            item["Editor"] = editor;

                            // January 5, 2015 8:53:42 PM	first.lastname@tenantad.forest
                            //item["Modified"] = DateTime.Parse("1/5/2015 8:53:42 PM");
                            //item["Editor"] = new FieldUserValue() { LookupId = user.Id };
                            item.SystemUpdate();
                            this.ClientContext.ExecuteQueryRetry();



                            item = onlineLibrary.GetItemById(ItemId);
                            this.ClientContext.Load(item);
                            this.ClientContext.ExecuteQueryRetry();


                            docId = item.RetrieveListItemValue("_dlc_DocId");
                            docIdUrl = item.RetrieveListItemValueAsHyperlink("_dlc_DocIdUrl");
                            modified = item.RetrieveListItemValue("Modified");
                            editor = item.RetrieveListItemUserValue("Editor");
                            redirectEmbeddedUrl = item.RetrieveListItemValue("ServerRedirectedEmbedUri");
                            fileRef = item.RetrieveListItemValue("FileRef");
                            title = item.RetrieveListItemValue("Title");
                            LogVerbose("[POST UPDATE] ==> DocId {0}  DocIdUrl {1}, Modified {2} Editor {3}, Embedded Url {4}, FileRef {5}", docId, (docIdUrl != null ? docIdUrl.Url : ""), modified, editor.Email, redirectEmbeddedUrl, fileRef);
                        }
                        else
                        {
                            LogWarning("Item ID {0} has a url which does not need to be updated.", ItemId);
                        }
                    }
                }
                catch (Exception ex)
                {
                    LogWarning("Failed to process list item {0} with message {1}", ItemId, ex.Message);
                }
            }
        }
    }
}
