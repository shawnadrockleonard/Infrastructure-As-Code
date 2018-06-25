using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Powershell.Commands.Development
{
    using InfrastructureAsCode.Powershell.CmdLets;
    using OfficeDevPnP.Core.AppModelExtensions;
    using OfficeDevPnP.Core.Extensions;
    using OfficeDevPnP.Core.Utilities;
    using Microsoft.SharePoint.Client;

    /// <summary>
    /// Enables Development Site Collection capabilities to deploy add-ins
    /// </summary>
    /// <remarks>
    ///     You must be a Tenant Administrator to enable this feature on a site collection
    /// </remarks>
    [Cmdlet(VerbsCommon.Add, "IaCEnableSideLoading")]
    public class AddIaCEnableSideLoading : IaCCmdlet
    {
        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();

            LogVerbose("Enables SharePoint app sideLoading");

            var sideLoadingGuid = OfficeDevPnP.Core.Constants.FeatureId_Site_AppSideLoading;
            var siteurl = this.ClientContext.Url;
            var outfilepath = siteurl.Replace(':', '_').Replace('/', '_');

            try
            {
                var site = this.ClientContext.Site;
                this.ClientContext.Load(site);

                var sideLoadingEnabled = AppCatalog.IsAppSideloadingEnabled(this.ClientContext);
                this.ClientContext.ExecuteQueryRetry();

                if (!sideLoadingEnabled.Value)
                {
                    LogVerbose("SideLoading feature is not enabled on the site: {0}", siteurl);

                    var siteFeatures = ClientContext.LoadQuery(ClientContext.Site.Features.Include(fctx => fctx.DefinitionId, fctx => fctx.DisplayName));
                    ClientContext.ExecuteQueryRetry();
                    var sideLoadingFeature = siteFeatures.FirstOrDefault(f => f.DefinitionId == sideLoadingGuid);


                    var siteFeature = site.Features.GetById(sideLoadingGuid);
                    this.ClientContext.Load(siteFeature);
                    this.ClientContext.ExecuteQueryRetry();


                    if (!siteFeature.ServerObjectIsNull())
                    {
                        LogWarning("Side loading feature is found.");
                    }

                    site.ActivateFeature(sideLoadingGuid, pollingIntervalSeconds: 20);
                    this.ClientContext.ExecuteQueryRetry();
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
