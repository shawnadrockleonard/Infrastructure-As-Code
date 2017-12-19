using InfrastructureAsCode.Powershell;
using InfrastructureAsCode.Powershell.CmdLets;
using InfrastructureAsCode.Core.Models;
using InfrastructureAsCode.Core.Extensions;
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
using System.Xml.Linq;
using InfrastructureAsCode.Powershell.PipeBinds;
using InfrastructureAsCode.Core.Utilities;

namespace InfrastructureAsCode.Powershell.Commands.ETL
{
    /// <summary>
    /// The function cmdlet will update a view
    /// </summary>
    [Cmdlet(VerbsCommon.Set, "IaCProvisionViews", SupportsShouldProcess = true)]
    [CmdletHelp("Set view definition based on JSON file.", Category = "ETL")]
    public class SetIaCProvisionViews : IaCCmdlet
    {
        /// <summary>
        /// Represents the directory path for any JSON files for serialization
        /// </summary>
        [Parameter(Mandatory = true, HelpMessage = "Provide a full path to the provisioner JSON file", Position = 0, ValueFromPipeline = true)]
        public string ProvisionerFilePath { get; set; }

        /// <summary>
        /// Specific list to be updated from the above action list
        /// </summary>
        [Parameter(Mandatory = true, ParameterSetName = "ActionDependency")]
        public ListPipeBind SpecificListName { get; set; }

        /// <summary>
        /// Specific view to be updated from the above action list
        /// </summary>
        [Parameter(Mandatory = true, ParameterSetName = "ActionDependency")]
        public ViewPipeBind SpecificViewName { get; set; }


        /// <summary>
        /// Validate parameters
        /// </summary>
        protected override void OnBeginInitialize()
        {
            if (!System.IO.File.Exists(this.ProvisionerFilePath))
            {
                var fileinfo = new System.IO.FileInfo(ProvisionerFilePath);
                throw new System.IO.FileNotFoundException("The provisioner file was not found", fileinfo.Name);
            }
        }

        /// <summary>
        /// Process the request
        /// </summary>
        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();



            // Retreive JSON Provisioner file and deserialize it
            var filePath = new System.IO.FileInfo(ProvisionerFilePath);
            var siteDefinition = JsonConvert.DeserializeObject<SiteProvisionerModel>(System.IO.File.ReadAllText(filePath.FullName));


            var listToUpdate = SpecificListName.GetList(this.ClientContext.Web);
            var viewToUpdate = SpecificViewName.GetView(listToUpdate);


            // Assumption here we are passing in existing list/view model
            var modelList = siteDefinition.Lists.FirstOrDefault(a => a.ListName.Equals(listToUpdate.Title, StringComparison.CurrentCultureIgnoreCase));
            var modelView = modelList.Views.FirstOrDefault(a => a.Title.Equals(viewToUpdate.Title, StringComparison.CurrentCultureIgnoreCase));
            var modelInternalView = modelList.InternalViews.FirstOrDefault(a => a.Title.Equals(viewToUpdate.Title, StringComparison.CurrentCultureIgnoreCase));


            if (modelView == null && modelInternalView == null)
            {
                LogWarning("Please select a valid view to modify");
                return;
            }

            var listName = this.SpecificListName;
            var context = this.ClientContext;
            var web = context.Web;
            var views = context.LoadQuery(listToUpdate.Views.Where(w => w.Title == viewToUpdate.Title).Include(tv => tv.Id,
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

            context.Load(web);
            context.Load(listToUpdate, ltu => ltu.Id, ltu => ltu.Title, ltut => ltut.RootFolder.ServerRelativeUrl);
            context.ExecuteQueryRetry();

            string webRelativeUrl = web.ServerRelativeUrl;


            if (modelView != null)
            {
                var thisview = views.FirstOrDefault(f => f.Title == viewToUpdate.Title);
                if (thisview != null)
                {
                    Guid viewID = thisview.Id;

                    var currentFields = thisview.ViewFields;
                    var newFields = modelView.FieldRefName;
                    currentFields.RemoveAll();
                    newFields.ToList().ForEach(vField =>
                    {
                        currentFields.Add(vField);
                    });

                    if (modelView.HasJsLink)
                    {
                        thisview.JSLink = modelView.JsLink;
                    }
                    thisview.RowLimit = modelView.RowLimit;
                    thisview.ViewQuery = modelView.ViewQuery;
                    thisview.Toolbar = string.Format("<Toolbar Type=\"{0}\"/>", modelView.ToolBarType.ToString());

                    if (this.ShouldProcess(string.Format("Should update view {0}", this.SpecificViewName)))
                    {
                        thisview.Update();
                        context.ExecuteQueryRetry();
                    }
                }
                else
                {
                    var view = listToUpdate.CreateView(modelView.InternalName, modelView.ViewCamlType, modelView.FieldRefName.ToArray(), modelView.RowLimit, modelView.DefaultView, modelView.ViewQuery, modelView.PersonalView, modelView.Paged);
                    context.Load(view, v => v.Title, v => v.Id, v => v.ServerRelativeUrl);
                    context.ExecuteQueryRetry();

                    view.Title = modelView.Title;
                    view.Toolbar = string.Format("<Toolbar Type=\"{0}\"/>", modelView.ToolBarType.ToString());
                    if (modelView.HasJsLink)
                    {
                        view.JSLink = modelView.JsLink;
                    }
                    view.Update();
                    context.ExecuteQueryRetry();
                }
            }
            else if (modelInternalView != null)
            {
                var weburl = TokenHelper.EnsureTrailingSlash(this.ClientContext.Web.ServerRelativeUrl);
                var listpageurl = string.Format("{0}{1}", weburl, modelInternalView.SitePage);
                var thisview = views.FirstOrDefault(f => f.ServerRelativeUrl == listpageurl);
                if (thisview != null)
                {
                    Guid viewID = thisview.Id;

                    var currentFields = thisview.ViewFields;
                    var newFields = modelInternalView.FieldRefName;
                    currentFields.RemoveAll();
                    newFields.ToList().ForEach(vField =>
                    {
                        currentFields.Add(vField);
                    });

                    if (modelInternalView.HasJsLink)
                    {
                        thisview.JSLink = modelInternalView.JsLink;
                    }
                    thisview.RowLimit = modelInternalView.RowLimit;
                    thisview.ViewQuery = modelInternalView.ViewQuery;
                    thisview.Toolbar = string.Format("<Toolbar Type=\"{0}\"/>", modelInternalView.ToolBarType.ToString());

                    if (this.ShouldProcess(string.Format("Should update view {0}", this.SpecificViewName)))
                    {
                        thisview.Update();
                        context.ExecuteQueryRetry();
                    }
                }
            }
        }

    }
}
