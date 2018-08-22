using InfrastructureAsCode.Core.Extensions;
using InfrastructureAsCode.Core.Models;
using InfrastructureAsCode.Powershell.Commands.Base;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Powershell.Commands.Sites
{
    /// <summary>
    /// The function cmdlet will retrieve site/web custom actions and remove the specified custom action by name
    /// </summary>
    [Cmdlet(VerbsCommon.Remove, "IaCCustomAction", SupportsShouldProcess = true)]
    public class RemoveIaCCustomAction : IaCCmdlet
    {
        [Parameter(Mandatory = true)]
        public string ActionName { get; set; }

        /// <summary>
        /// Execute the removal if found
        /// </summary>
        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();

            var site = this.ClientContext.Site;
            var web = this.ClientContext.Web;
            this.ClientContext.Load(site, ccsu => ccsu.ServerRelativeUrl, cssu => cssu.UserCustomActions);
            this.ClientContext.Load(web, ccwu => ccwu.ServerRelativeUrl, ccwu => ccwu.UserCustomActions);
            this.ClientContext.ExecuteQueryRetry();


            if (site.CustomActionExists(ActionName))
            {
                var customActions = site.UserCustomActions.Where(fn => fn.Name == ActionName);
                var actionIds = new List<Guid>();
                actionIds.AddRange(customActions.Select(s => s.Id));

                foreach (var customActionId in actionIds)
                {
                    site.DeleteCustomAction(customActionId);
                }
            }

            if (web.CustomActionExists(ActionName))
            {
                var customActions = web.UserCustomActions.Where(fn => fn.Name == ActionName);
                var actionIds = new List<Guid>();
                actionIds.AddRange(customActions.Select(s => s.Id));

                foreach (var customActionId in actionIds)
                {
                    web.DeleteCustomAction(customActionId);
                }
            }



            if (site.RemoveCustomActionLink(ActionName))
            {
                LogVerbose("Successfully removed [Site] action {0}", ActionName);
            }


            if (web.RemoveCustomActionLink(ActionName))
            {
                LogVerbose("Successfully removed [Web] action {0}", ActionName);
            }

        }
    }
}
