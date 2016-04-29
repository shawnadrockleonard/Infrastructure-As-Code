using IaC.Powershell.CmdLets;
using IaC.Core.Models;
using IaC.Core.Extensions;
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

namespace IaC.Powershell.Commands.Lists
{
    /// <summary>
    /// The function cmdlet will create a site page and add a webpart with a specific hidden list definition
    /// </summary>
    [Cmdlet(VerbsCommon.Set, "IaCWebPartDefinition")]
    [CmdletHelp("Identify users via json file and send email", Category = "Lists")]
    public class SetIaCWebPartDefinition : IaCCmdlet
    {
        /// <summary>
        /// The location where pages will be created
        /// </summary>
        private string SitePagesLibraryTitle = "Site Pages";
        /// <summary>
        /// The relative URI to the library
        /// </summary>
        private string SitePagesRelativeUrl = "SitePages";

        [Parameter(Mandatory = true)]
        public string ListTitle { get; set; }

        [Parameter(Mandatory = true)]
        public string ViewTitle { get; set; }

        [Parameter(Mandatory = true)]
        public string PageTitle { get; set; }

        [Parameter(Mandatory = true)]
        public string WebPartTitle { get; set; }

        [Parameter(Mandatory = true)]
        public string WebPartViewTitle { get; set; }

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

            string pageFileName = PageTitle.Replace(" ", string.Empty);
            if(pageFileName.IndexOf("aspx") <= -1)
            {
                pageFileName = String.Format("{0}.aspx", pageFileName);
            }
           


            var context = this.ClientContext;
            var web = context.Web;
            context.Load(web);
            context.ExecuteQuery();
            string webRelativeUrl = web.ServerRelativeUrl;

            ListCollection allLists = web.Lists;
            IEnumerable<List> foundLists = context.LoadQuery(allLists.Where(list => list.Title == ListTitle)
                .Include(wl => wl.Id, wl => wl.Title, wl => wl.RootFolder.ServerRelativeUrl));
            context.ExecuteQuery();
            List webPartList = foundLists.FirstOrDefault();

            var wikiPageUrl = web.AddWikiPage(SitePagesLibraryTitle, pageFileName);
            context.ExecuteQuery();
            if (string.IsNullOrEmpty(wikiPageUrl))
            {
                wikiPageUrl = string.Format("{0}/{1}", SitePagesRelativeUrl, pageFileName);
            }

            var listUrl = webPartList.RootFolder.ServerRelativeUrl;
            var pageUrl = string.Format("{0}/{1}", webRelativeUrl, wikiPageUrl);
  

            var views = webPartList.Views;
            ClientContext.Load(views);
            ClientContext.ExecuteQueryRetry();

            var viewFound = false;
            foreach (var view in views.Where(w => w.Hidden))
            {
                LogVerbose("View {0} with title {1} and hidden status:{2} and server relative url:{3}", view.Id, view.Title, view.Hidden, view.ServerRelativeUrl);
                if (view.ServerRelativeUrl.ToUpper() == pageUrl.ToUpper())
                {
                    viewFound = true;
                }
            }

            var thisview = webPartList.GetViewByName(WebPartViewTitle);
            webPartList.Context.Load(thisview, tv => tv.Id, tv => tv.ServerRelativeUrl, tv => tv.HtmlSchemaXml, tv => tv.RowLimit, tv => tv.JSLink, tv => tv.ViewFields, tv => tv.ViewQuery, tv => tv.Aggregations);
            webPartList.Context.ExecuteQueryRetry();
            Guid viewID = thisview.Id;

            if (!viewFound)
            {
                // change layout to header and two column
                web.AddLayoutToWikiPage(SitePagesRelativeUrl, OfficeDevPnP.Core.WikiPageLayout.OneColumn, pageFileName);
                web.AddHtmlToWikiPage(SitePagesRelativeUrl, "<div>This is a provisioning TEST!!!</div>", pageFileName, 1, 1);

                // get markup and place on page
                var xlstMarkup = webPartList.GetXsltWebPartXML(pageUrl, WebPartTitle, viewID);
                WebPartEntity wpe = new WebPartEntity();
                wpe.WebPartXml = xlstMarkup;
                wpe.WebPartTitle = WebPartTitle;
                wpe.WebPartIndex = 1;
                web.AddWebPartToWikiPage(SitePagesRelativeUrl, wpe, pageFileName, 1, 1, true, false);

                // Hidden View was created for this
                views = webPartList.Views;
                ClientContext.Load(views);
                ClientContext.ExecuteQueryRetry();
            }

            // Get View Markup
            var viewHtml = thisview.HtmlSchemaXml;
            var viewXdoc = XDocument.Parse(viewHtml);
            foreach (var view in views.Where(w => w.Hidden && w.ServerRelativeUrl == pageUrl))
            {
                LogVerbose("Found View {0} with title {1} and hidden status:{2} and server relative url:{3}", view.Id, view.Title, view.Hidden, view.ServerRelativeUrl);
                var vview = viewXdoc.Root.Element("Query");
                var toolbar = viewXdoc.Root.Element("Toolbar");
                var aggregate = viewXdoc.Root.Element("Aggregations");
                if (aggregate != null)
                {
                    view.Aggregations = aggregate.ToString(SaveOptions.DisableFormatting);
                }

                view.ViewFields.RemoveAll();
                foreach (var vf in thisview.ViewFields)
                {
                    view.ViewFields.Add(vf);
                }
                view.RowLimit = thisview.RowLimit;
                view.Toolbar = "<Toolbar Type=\"None\"/>";
                view.ViewQuery = thisview.ViewQuery;
                view.JSLink = thisview.JSLink;
                view.Aggregations = thisview.Aggregations;

                view.Update();
                ClientContext.ExecuteQueryRetry();
            }
        }

    }
}
