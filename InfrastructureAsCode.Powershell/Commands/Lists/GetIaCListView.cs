using InfrastructureAsCode.Powershell.CmdLets;
using InfrastructureAsCode.Powershell.PipeBinds;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace InfrastructureAsCode.Powershell.Commands.Lists
{
    /// <summary>
    /// Returns one or all views from a list
    /// </summary>
    /// <remarks>
    /// Get-IaCListView -List ""Demo List""
    /// Get-IaCListView -List ""Demo List"" -Identity ""Demo View""
    /// Get-IaCListView -List ""Demo List"" -Identity ""5275148a-6c6c-43d8-999a-d2186989a661""
    /// </remarks>
    [Cmdlet(VerbsCommon.Get, "IaCListView")]
    public class GetIaCListView : IaCCmdlet
    {
        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0, HelpMessage = "The ID or Url of the list.")]
        public ListPipeBind List { get; set; }

        [Parameter(Mandatory = false)]
        public ViewPipeBind Identity { get; set; }

        public override void ExecuteCmdlet()
        {

            if (List != null)
            {
                var SelectedWeb = this.ClientContext.Web;

                var list = List.GetList(SelectedWeb);
                if (list != null)
                {
                    View view = null;
                    IEnumerable<View> views = null;
                    if (Identity != null)
                    {
                        view = Identity.GetView(list);
                        if (view != null)
                        {
                            view.EnsureProperties(v => v.ViewFields, v => v.JSLink, v => v.ViewQuery);
                            var doc = XDocument.Parse(string.Format("<ViewXml>{0}</ViewXml>", view.ViewQuery));
                            string indented = doc.ToString();
                            LogVerbose("View {0} CAML:{1}", view.Title, indented); // write query and Title
                        }

                        WriteObject(view);
                    }
                    else
                    {
                        views = ClientContext.LoadQuery(list.Views.IncludeWithDefaultProperties(v => v.ViewFields));
                        ClientContext.ExecuteQueryRetry();
                        WriteObject(views, true);
                    }
                }
            }
        }
    }
}
