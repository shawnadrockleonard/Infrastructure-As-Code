using InfrastructureAsCode.Powershell.CmdLets;
using InfrastructureAsCode.Powershell.PipeBinds;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Powershell.Commands.Lists
{
    /// <summary>
    /// Returns a List object
    /// </summary>
    /// <remarks>
    /// Get-IaCList
    /// Get-IaCList -Identity /Lists/Announcements
    /// Get-IaCList -Identity 99a00f6e-fb81-4dc7-8eac-e09c6f9132fe"
    /// </remarks>
    [Cmdlet(VerbsCommon.Get, "IaCList")]
    public class GetIaCList : IaCCmdlet
    {
        [Parameter(Mandatory = false, ValueFromPipeline = true, Position = 0, HelpMessage = "The ID or Url of the list.")]
        public ListPipeBind Identity;

        public override void ExecuteCmdlet()
        {
            var SelectedWeb = this.ClientContext.Web;

            if (Identity != null)
            {
                var list = Identity.GetList(SelectedWeb);
                WriteObject(list);

            }
            else
            {
                var lists = ClientContext.LoadQuery(SelectedWeb.Lists.IncludeWithDefaultProperties(l => l.Id, l => l.BaseTemplate, l => l.OnQuickLaunch, l => l.DefaultViewUrl, l => l.Title, l => l.Hidden, l => l.RootFolder.ServerRelativeUrl));
                ClientContext.ExecuteQueryRetry();
                WriteObject(lists, true);
            }
        }
    }
}