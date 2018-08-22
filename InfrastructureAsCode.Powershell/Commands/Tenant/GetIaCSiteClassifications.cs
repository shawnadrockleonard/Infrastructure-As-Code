using InfrastructureAsCode.Powershell.Commands.Base;
using InfrastructureAsCode.Core.HttpServices;
using InfrastructureAsCode.Core.Extensions;
using Microsoft.Online.SharePoint.TenantAdministration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using InfrastructureAsCode.Core.oAuth;
using System.Threading.Tasks;
using OfficeDevPnP.Core.Framework.Graph;

namespace InfrastructureAsCode.Powershell.Commands.Tenant
{
    /// <summary>
    /// Query the tenant for the enabled Classifications
    /// </summary>
    [Cmdlet(VerbsCommon.Get, "IaCSiteClassifications")]
    public class GetIaCSiteClassifications : IaCAdminCmdlet
    {
        #region Parameters

        [Parameter(Mandatory = true)]
        public string TenantId { get; set; }

        [Parameter(Mandatory = true)]
        public string MSALClientID { get; set; }

        [Parameter(Mandatory = true)]
        public string MSALClientSecret { get; set; }

        [Parameter(Mandatory = true)]
        public string PostLogoutRedirectURI { get; set; }

        #endregion


        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();


            var _list = new List<string>();
            var tenantAdminUri = ClientContext.GetTenantAdminUri();
            var tenantRealmId = (new Uri(tenantAdminUri)).GetRealmFromTargetUrl();

            var appConfig = new AppConfig()
            {
                TenantId = TenantId,
                MSALClientID = MSALClientID,
                MSALClientSecret = MSALClientSecret,
                PostLogoutRedirectURI = PostLogoutRedirectURI
            };

            LogVerbose($"Parameter TenantId {TenantId} and returned RealmId from ClientContext {tenantRealmId}");

            var _graphClient = new GraphHttpHelper(appConfig);
            var _accessTokenTask = Task.Run(async () => await _graphClient.GetGraphDaemonAccessTokenAsync());

            try
            {
                System.Diagnostics.Trace.TraceInformation($"Calling EnsureToken {tenantAdminUri}");
                _accessTokenTask.Wait();

                LogVerbose($"Calling Tenant Directory Settings {tenantAdminUri}");
                var _classifications = SiteClassificationsUtility.GetSiteClassificationsSettings(_accessTokenTask.Result);
                LogVerbose($"Classification Default {_classifications.DefaultClassification}");
                LogVerbose($"Classification Private {_classifications.UsageGuidelinesUrl}");

                _list.AddRange(_classifications.Classifications);
            }
            catch (Exception ex)
            {
                LogError(ex, $"Failed to retreive tenant classifications {ex.Message}");
            }

            WriteObject(_list);
        }
    }
}
