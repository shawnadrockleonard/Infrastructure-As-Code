using InfrastructureAsCode.Powershell.Commands.Base;
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
    /// <summary>
    /// sets the user as an admin
    /// </summary>
    [Cmdlet(VerbsCommon.Set, "IaCSiteAdmin")]
    [CmdletHelp("sets the user as an admin", Category = "Principals")]
    public class SetIaCSiteAdmin : IaCAdminCmdlet
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
        /// If specified it will add the user as an administrator
        /// </summary>
        [Parameter(Mandatory = false, HelpMessage = "Provides a specific site to query and manipulate")]
        public SwitchParameter IsAdmin { get; set; }


        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();

            SetSiteAdmin(this.SiteUrl, UserName, IsAdmin.ToBool());
        }

    }
}
