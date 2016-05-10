using InfrastructureAsCode.Powershell.CmdLets;
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Powershell.Commands.Principals
{
    [Cmdlet(VerbsCommon.Set, "IaCSiteAdmin")]
    [CmdletHelp("sets the user as an admin", Category = "Principals")]
    public class SetIaCSiteAdmin : SPOAdminCmdlet
    {
        /// <summary>
        /// The absolute URL to the site collection or web
        /// </summary>
        [Parameter(Mandatory = true, HelpMessage = "Provides a specific site to query and manipulate")]
        public string SiteUrl { get; set; }

        /// <summary>
        /// The absolute URL to the site collection or web
        /// </summary>
        [Parameter(Mandatory = false, HelpMessage = "username to grant/revoke admin")]
        public string UserName { get; set; }

        /// <summary>
        /// The absolute URL to the site collection or web
        /// </summary>
        [Parameter(Mandatory = false, HelpMessage = "Provides a specific site to query and manipulate")]
        public bool IsAdmin { get; set; }


        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();

            SetSiteAdmin(this.SiteUrl, UserName, IsAdmin);
        }

    }
}
