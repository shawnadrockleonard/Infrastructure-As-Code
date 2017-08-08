using InfrastructureAsCode.Core.Models.Enums;
using InfrastructureAsCode.Powershell.CmdLets;
using InfrastructureAsCode.Powershell.PipeBinds;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;

namespace InfrastructureAsCode.Powershell.Commands.Lists
{

    [Cmdlet(VerbsCommon.Set, "IaCListView", SupportsShouldProcess = false)]
    public class SetIaCListView : IaCCmdlet
    {
        /// <summary>
        /// Internal Names of the View
        /// </summary>
        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
        public ListPipeBind List { get; set; }

        /// <summary>
        /// Internal Names of the View
        /// </summary>
        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 1)]
        public ViewPipeBind Identity { get; set; }

        /// <summary>
        /// Internal Names of the View
        /// </summary>
        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 2)]
        public string QueryXml { get; set; }

        /// <summary>
        /// Internal Names of the View
        /// </summary>
        [Parameter(Mandatory = false, ValueFromPipeline = true, Position = 3)]
        public string[] ViewFields { get; set; }

        /// <summary>
        /// Internal Names of the View
        /// </summary>
        [Parameter(Mandatory = false, ValueFromPipeline = true, Position = 4)]
        public Nullable<int> RowLimit { get; set; }

        /// <summary>
        /// Internal Names of the View
        /// </summary>
        [Parameter(Mandatory = false, ValueFromPipeline = true, Position = 5)]
        public string[] JsLinkUris { get; set; }

        /// <summary>
        /// Internal Names of the View
        /// </summary>
        [Parameter(Mandatory = false, ValueFromPipeline = true, Position = 6)]
        public Nullable<ViewToolBarEnum> Toolbar { get; set; }

        /// <summary>
        /// Internal Names of the View
        /// </summary>
        [Parameter(Mandatory = false, HelpMessage = "If specified, the view will have paging.")]
        public SwitchParameter PagedView { get; set; }

        [Parameter(Mandatory = false, HelpMessage = "If specified, a personal view will be created.")]
        public SwitchParameter Personal;

        [Parameter(Mandatory = false, HelpMessage = "If specified, the view will be set as the default view for the list.")]
        public SwitchParameter SetAsDefault;


        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();



            var listName = List.Title;
            var context = this.ClientContext;
            var web = context.Web;
            context.Load(web);
            context.ExecuteQueryRetry();
            string webRelativeUrl = web.ServerRelativeUrl;
            ListCollection allLists = web.Lists;
            IEnumerable<List> foundLists = context.LoadQuery(allLists.Where(list => list.Title == listName).Include(wl => wl.Id, wl => wl.Title, wl => wl.RootFolder.ServerRelativeUrl));
            context.ExecuteQueryRetry();



            var viewName = Identity.Title;
            var internalRowLimit = Convert.ToUInt32(RowLimit.HasValue ? RowLimit.Value : 50);

            List listToUpdate = foundLists.FirstOrDefault();
            var views = listToUpdate.Views;
            ClientContext.Load(views, v => v.Include(tv => tv.Id,
                tv => tv.Title,
                tv => tv.ServerRelativeUrl,
                tv => tv.DefaultView,
                tv => tv.HtmlSchemaXml,
                tv => tv.RowLimit,
                tv => tv.Toolbar,
                tv => tv.JSLink,
                tv => tv.ViewFields,
                tv => tv.ViewQuery,
                tv => tv.Aggregations,
                tv => tv.AggregationsStatus,
                tv => tv.Hidden,
                tv => tv.Method,
                tv => tv.PersonalView,
                tv => tv.ReadOnlyView,
                tv => tv.ViewType));
            ClientContext.ExecuteQueryRetry();


            var thisview = Identity.GetView(listToUpdate);
            if (thisview != null)
            {
                Guid viewID = thisview.Id;


                if (ViewFields != null && ViewFields.Length > 0)
                {
                    var currentFields = thisview.ViewFields;
                    currentFields.RemoveAll();
                    ViewFields.ToList().ForEach(vField =>
                    {
                        currentFields.Add(vField.Trim());
                    });
                }

                if (JsLinkUris != null && JsLinkUris.Count() > 0)
                {
                    var jsLinkUri = String.Join("|", JsLinkUris);
                    thisview.JSLink = jsLinkUri;
                }
                thisview.RowLimit = internalRowLimit;
                thisview.ViewQuery = QueryXml;

                if (Toolbar.HasValue)
                {
                    thisview.Toolbar = string.Format("<Toolbar Type=\"{0}\"/>", Toolbar.ToString());
                }

                if (this.ShouldProcess(string.Format("Should update view {0}", viewName)))
                {
                    thisview.Update();
                    listToUpdate.Context.ExecuteQueryRetry();
                }
            }
            else
            {
                var internalName = viewName.Replace(" ", string.Empty);

                if (this.ShouldProcess(string.Format("Should create the view {0}", viewName)))
                {
                    var view = listToUpdate.CreateView(internalName, ViewType.None, ViewFields, internalRowLimit, SetAsDefault, QueryXml, Personal, PagedView);
                    listToUpdate.Context.Load(view, v => v.Title, v => v.Id, v => v.ServerRelativeUrl);
                    listToUpdate.Context.ExecuteQueryRetry();

                    view.Title = viewName;

                    if (Toolbar.HasValue)
                    {
                        thisview.Toolbar = string.Format("<Toolbar Type=\"{0}\"/>", Toolbar.ToString());
                    }

                    if (JsLinkUris != null && JsLinkUris.Count() > 0)
                    {
                        var jsLinkUri = String.Join("|", JsLinkUris);
                        thisview.JSLink = jsLinkUri;
                    }

                    view.Update();
                    listToUpdate.Context.ExecuteQueryRetry();
                }
            }

        }
    }
}
