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

namespace InfrastructureAsCode.Powershell.Commands.Lists
{
    /// <summary>
    /// The function cmdlet will update view definitions
    /// </summary>
    [Cmdlet(VerbsCommon.Set, "IaCViewDefinition")]
    [CmdletHelp("Update a view definition", Category = "Lists")]
    public class SetIaCViewDefinition : IaCCmdlet
    {
        /// <summary>
        /// Represents the title of the list being updated
        /// </summary>
        [Parameter(Mandatory = true)]
        public string ListTitle { get; set; }

        /// <summary>
        /// Represents the title of the view being updated
        /// </summary>
        [Parameter(Mandatory = true)]
        public string ViewTitle { get; set; }

        /// <summary>
        /// Represents the caml query for the view being updated
        /// </summary>
        [Parameter(Mandatory = true)]
        public string CamlQuery { get; set; }

        /// <summary>
        /// A string array of JsLinks to augment the view
        /// </summary>
        [Parameter(Mandatory = false)]
        public string[] JsLinkUris { get; set; }
        

        /// <summary>
        /// Process the request
        /// </summary>
        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();

            if (this.ClientContext == null)
            {
                LogWarning("Invalid client context, configure the service to run again");
                return;
            }

            var context = this.ClientContext;
            var web = context.Web;
            context.Load(web);
            context.ExecuteQuery();
            string webRelativeUrl = web.ServerRelativeUrl;
            ListCollection allLists = web.Lists;
            IEnumerable<List> foundLists = context.LoadQuery(allLists.Where(list => list.Title == ListTitle).Include(wl => wl.Id, wl => wl.Title, wl => wl.RootFolder.ServerRelativeUrl));
            context.ExecuteQuery();

            List accessRequest = foundLists.FirstOrDefault();
            var views = accessRequest.Views;
            ClientContext.Load(views, v => v.Include(tv => tv.Id, tv => tv.Title, tv => tv.ServerRelativeUrl, tv => tv.HtmlSchemaXml, tv => tv.RowLimit, tv => tv.JSLink, tv => tv.ViewFields, tv => tv.ViewQuery, tv => tv.Aggregations));
            ClientContext.ExecuteQueryRetry();

            var thisview = views.FirstOrDefault(f => f.Title.Equals(ViewTitle, StringComparison.CurrentCultureIgnoreCase));
            if (thisview != null)
            {
                Guid viewID = thisview.Id;

                if(JsLinkUris!=null && JsLinkUris.Count() > 0)
                {
                    var jsLinkUri = String.Join("|", JsLinkUris);
                    thisview.JSLink = jsLinkUri;
                }

                thisview.ViewQuery = CamlQuery;

                if (!DoNothing)
                {
                    thisview.Update();
                    context.ExecuteQueryRetry();
                }
            }
        }

    }
}
