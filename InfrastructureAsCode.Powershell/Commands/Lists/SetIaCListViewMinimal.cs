using InfrastructureAsCode.Powershell;
using InfrastructureAsCode.Powershell.CmdLets;
using InfrastructureAsCode.Core.Models;
using InfrastructureAsCode.Powershell.Extensions;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using OfficeDevPnP.Core.Entities;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using InfrastructureAsCode.Powershell.PipeBinds;

namespace InfrastructureAsCode.Powershell.Commands.Lists
{
    /// <summary>
    /// The function cmdlet will update view definitions
    /// </summary>
    /// <remarks>
    /// Set-IaCListViewMinimal -List "Sample List" -Identity "Sample View" -RowLimit 5 -Verbose
    /// Set-IaCListViewMinimal -List $list -Identity "Sample View" -RowLimit 5 -Verbose
    /// </remarks>
    [Cmdlet(VerbsCommon.Set, "IaCListViewMinimal", SupportsShouldProcess = true)]
    [CmdletHelp("Update a view definition", Category = "Lists")]
    public class SetIaCListViewMinimal : IaCCmdlet
    {
        /// <summary>
        /// Represents the title of the list being updated
        /// </summary>
        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 1)]
        public ListPipeBind List { get; set; }

        /// <summary>
        /// Represents the title of the view being updated
        /// </summary>
        [Parameter(Mandatory = true, ValueFromPipeline = true, Position =2)]
        public ViewPipeBind Identity { get; set; }

        /// <summary>
        /// Represents the caml query for the view being updated
        /// </summary>
        [Parameter(Mandatory = false, ValueFromPipeline = true, Position = 3)]
        public string CamlQuery { get; set; }

        /// <summary>
        /// A string array of JsLinks to augment the view
        /// </summary>
        [Parameter(Mandatory = false, ValueFromPipeline = true, Position = 4)]
        public string[] JsLinkUris { get; set; }

        /// <summary>
        /// Internal Names of the View
        /// </summary>
        [Parameter(Mandatory = false, ValueFromPipeline = true, Position = 5)]
        public Nullable<uint> RowLimit { get; set; }


        /// <summary>
        /// Process the request
        /// </summary>
        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();


            var context = this.ClientContext;
            var web = context.Web;
            context.Load(web);
            context.ExecuteQueryRetry();


            var listDefinition = List.GetList(web);
            context.Load(listDefinition,
                wl => wl.Id, wl => wl.Title, wl => wl.RootFolder.ServerRelativeUrl,
                wl => wl.Views.Include(tv => tv.Id, tv => tv.Title, tv => tv.ServerRelativeUrl, tv => tv.HtmlSchemaXml, tv => tv.RowLimit, tv => tv.JSLink, tv => tv.ViewFields, tv => tv.ViewQuery, tv => tv.Aggregations));
            context.ExecuteQueryRetry();


            var thisview = Identity.GetView(listDefinition);
            if (thisview != null)
            {
                Guid viewID = thisview.Id;

                if (JsLinkUris != null && JsLinkUris.Count() > 0)
                {
                    var jsLinkUri = String.Join("|", JsLinkUris);
                    thisview.JSLink = jsLinkUri;
                }

                if(RowLimit.HasValue)
                {
                    thisview.RowLimit = RowLimit.Value;
                }

                if (!string.IsNullOrEmpty(CamlQuery))
                {
                    thisview.ViewQuery = CamlQuery;
                }

                if (this.ShouldProcess(string.Format("Updating {0} view", thisview.Title)))
                {
                    thisview.Update();
                    context.ExecuteQueryRetry();
                }
            }
        }

    }
}
