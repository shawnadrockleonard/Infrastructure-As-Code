using InfrastructureAsCode.Core;
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
    /// This command will enable or disable app side loading for the site
    /// </summary>
    [Cmdlet(VerbsCommon.Set, "IaCAppSideLoading")]
    [CmdletHelp("Will enable or disable app side loading", Category = "Development")]
    public class SetIaCAppSideLoading : IaCCmdlet
    {
        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();

            LogVerbose("To enable SharePoint app sideLoading, enter Site Url, username and password");

            var sideLoadingGuid = OfficeDevPnP.Core.Constants.APPSIDELOADINGFEATUREID;
            var siteurl = this.ClientContext.Url;
            var outfilepath = siteurl.Replace(':', '_').Replace('/', '_');

            try
            {
                var web = this.ClientContext.Web;
                var site = this.ClientContext.Site;
                this.ClientContext.Load(web);
                this.ClientContext.Load(site);
                this.ClientContext.ExecuteQuery();

                var siteFeatures = site.Features;
                this.ClientContext.Load(siteFeatures, fcol => fcol.Include(f => f.DisplayName, f => f.DefinitionId));
                var webFeatures = web.Features;
                this.ClientContext.Load(webFeatures, wf => wf.Include(f => f.DisplayName, f => f.DefinitionId));
                this.ClientContext.ExecuteQuery();

                LogWarning("Now parsing site features.");

                siteFeatures.ToList().ForEach(a =>
                {
                    var status = FeatureExtensions.IsFeatureActive(site, a.DefinitionId);

                    LogVerbose("Site Feature {0} with Id {1} and status {2}", a.DisplayName, a.DefinitionId, status);
                });

                LogWarning("Now parsing web features.");

                webFeatures.ToList().ForEach(a =>
                {
                    var status = FeatureExtensions.IsFeatureActive(web, a.DefinitionId);

                    LogVerbose("Web Feature {0} with Id {1} and status {2}", a.DisplayName, a.DefinitionId, status);
                });


                var sideLoadingEnabled = AppCatalog.IsAppSideloadingEnabled(this.ClientContext);
                this.ClientContext.ExecuteQuery();

                var sideLoadingFeature = siteFeatures.FirstOrDefault(f => f.DefinitionId == sideLoadingGuid);

                if (!sideLoadingEnabled.Value)
                {
                    LogVerbose("SideLoading feature is not enabled on the site: {0}", siteurl);
                    var siteFeature = site.Features.GetById(sideLoadingGuid);
                    this.ClientContext.Load(siteFeature);
                    this.ClientContext.ExecuteQuery();

                    if (siteFeature != null)
                    {
                        LogWarning("Side loading feature is found.");
                    }

                    site.Features.Add(sideLoadingGuid, true, Microsoft.SharePoint.Client.FeatureDefinitionScope.Site);
                    this.ClientContext.ExecuteQuery();
                    LogVerbose("SideLoading feature enabled on site {0}", siteurl);
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
