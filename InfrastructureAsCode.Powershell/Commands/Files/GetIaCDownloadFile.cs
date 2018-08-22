using InfrastructureAsCode.Powershell.Commands.Base;
using InfrastructureAsCode.Powershell.PipeBinds;
using InfrastructureAsCode.Core.Extensions;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;


namespace InfrastructureAsCode.Powershell.Commands.Files
{
    /// <summary>
    /// Downloads a file from the ListItem
    /// </summary>
    /// <remarks>
    /// Get-IaCDownloadFile -List ""Demo List""
    /// </remarks>
    [Cmdlet(VerbsCommon.Get, "IaCDownloadFile")]
    public class GetIaCDownloadFile : IaCCmdlet
    {
        /// <summary>
        /// The list from which the listitem will be queried
        /// </summary>
        [Parameter(Mandatory = true, Position = 0)]
        public ListPipeBind List { get; set; }

        /// <summary>
        /// The unique ID of the file to be downloaded
        /// </summary>
        [Parameter(Mandatory = true, Position = 1)]
        public int ItemId { get; set; }

        /// <summary>
        /// 
        /// </summary>
        [Parameter(Mandatory = false, Position = 2)]
        public string Path { get; set; }

        public override void ExecuteCmdlet()
        {
            var PathIO = new System.IO.DirectoryInfo(Path);

            var SelectedWeb = this.ClientContext.Web;

            if (string.IsNullOrEmpty(Path) || !PathIO.Exists)
            {
                // build the directory structure in the users AppData for file download/upload
                string result = System.IO.Path.GetTempPath();
                var resultdir = new System.IO.DirectoryInfo(result);
                // create the path/directory and inherit permissions
                PathIO = resultdir.CreateSubdirectory("iaclogs", resultdir.GetAccessControl());
            }

            // Pull listitem from list
            var list = List.GetList(SelectedWeb);
            if (list != null)
            {
                // Commented out to demonstrate extension handling URL property initialization
                // ClientContext.Load(SelectedWeb, wctx => wctx.Url);

                var item = list.GetItemById(ItemId);
                ClientContext.Load(item,
                    ictx => ictx.Id,
                    ictx => ictx["Title"],
                    ictx => ictx["FileLeafRef"]);
                ClientContext.ExecuteQueryRetry();


                var filepath = item.DownloadFile(ClientContext, PathIO.FullName);
                LogVerbose("Successfully downloaded {0}", filepath);
            }

        }
    }
}
