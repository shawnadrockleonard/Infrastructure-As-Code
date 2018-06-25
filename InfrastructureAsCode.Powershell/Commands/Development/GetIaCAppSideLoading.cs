using InfrastructureAsCode.Core;
using InfrastructureAsCode.Core.Extensions;
using InfrastructureAsCode.Powershell.CmdLets;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.AppModelExtensions;
using OfficeDevPnP.Core.Extensions;
using OfficeDevPnP.Core.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;
using Resources = InfrastructureAsCode.Core.Properties.Resources;

namespace InfrastructureAsCode.Powershell.Commands.Development
{
    /// <summary>
    /// This command will check if side loading is enabled for the site
    /// </summary>
    [Cmdlet(VerbsCommon.Get, "IaCAppSideLoading")]
    [CmdletHelp("Will enable or disable app side loading", Category = "Development")]
    public class GetIaCAppSideLoading : IaCCmdlet
    {
        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();

            var siteurl = this.ClientContext.Url.EnsureTrailingSlashLowered();
            LogVerbose($"Check if app sideLoading is enabled for {siteurl}");


            try
            {

                var sideLoadingEnabled = AppCatalog.IsAppSideloadingEnabled(this.ClientContext);
                this.ClientContext.ExecuteQuery();

                if (!sideLoadingEnabled.Value)
                {
                    LogWarning("SideLoading feature is not enabled on the site: {0}", siteurl);
                }
                else
                {
                    LogVerbose("SideLoading feature is already enabled on site {0}", siteurl);
                }
            }
            catch (Exception ex)
            {
                LogError(ex, "Error encountered when trying to enable SideLoading feature {0} with message {1}", siteurl, ex.Message.ToString());
            }

        }
    }
}
