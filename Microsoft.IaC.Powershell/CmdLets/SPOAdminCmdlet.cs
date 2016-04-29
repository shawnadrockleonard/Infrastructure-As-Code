using System;
using System.Management.Automation;
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.PowerShell.Commands;
using Microsoft.SharePoint.Client;
using IaC.Core.Enums;
using Resources = IaC.Core.Properties.Resources;
using Microsoft.Online.SharePoint.TenantManagement;

namespace IaC.Powershell.CmdLets
{
    /// <summary>
    /// SharePoint Online base command for tenant level administration
    /// </summary>
    public abstract class SPOAdminCmdlet :IaCCmdlet
    {
        private Tenant _tenant;
        /// <summary>
        /// Create client context to tenant admin
        /// </summary>
        public Tenant Tenant
        {
            get
            {
                if (_tenant == null)
                {
                    _tenant = new Tenant(this.ClientContext);
                }
                return _tenant;
            }
        }

        private Office365Tenant _officeTenant { get; set; }

        /// <summary>
        /// Initializes a Office365 Tenant Context
        /// </summary>
        public Office365Tenant OfficeTenant
        {
            get
            {
                if (_officeTenant == null)
                {
                    _officeTenant = new Office365Tenant(this.ClientContext);
                }
                return _officeTenant;
            }
        }

        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            if (SPIaCConnection.CurrentConnection == null)
            {
                throw new InvalidOperationException(Resources.NoConnection);
            }
            if (ClientContext == null)
            {
                throw new InvalidOperationException(Resources.NoConnection);
            }

            SPIaCConnection.CurrentConnection.CacheContext();

            Uri uri = new Uri(this.ClientContext.Url);
            var urlParts = uri.Authority.Split(new[] { '.' });
            if (!urlParts[0].EndsWith("-admin"))
            {
                var adminUrl = string.Format("https://{0}-admin.{1}.{2}", urlParts[0], urlParts[1], urlParts[2]);

                SPIaCConnection.CurrentConnection.Context = this.ClientContext.Clone(adminUrl);
            }

        }

        protected override void EndProcessing()
        {
            SPIaCConnection.CurrentConnection.RestoreCachedContext();
            if (Tenant.Context == null)
            {
                Tenant.Context.Dispose();
            }
        }

        /// <summary>
        /// Sets the site collection administrator for the activity
        /// </summary>
        /// <param name="_siteUrl">The relative url to the site collection</param>
        /// <param name="userNameWithoutClaims">Provide the username without the claim prefix</param>
        /// <param name="isSiteAdmin">(OPTIONAL) true to set the user as a site collection administrator</param>
        protected virtual void SetSiteAdmin(string _siteUrl, string userNameWithoutClaims, bool isSiteAdmin = false)
        {
            var claimProviderUserName = string.Format("i:0#.f|membership|{0}", userNameWithoutClaims);
            if (isSiteAdmin)
            {
                LogVerbose("Granting access to {0} for {1}", _siteUrl, claimProviderUserName);

            }
            else
            {
                LogVerbose("Removing access to {0} for {1}", _siteUrl, claimProviderUserName);
            }

            try
            {
                if (Tenant.Context == null)
                {
                    this.ClientContext.Load(Tenant);
                    this.ClientContext.ExecuteQuery();
                }
                Tenant.SetSiteAdmin(_siteUrl, claimProviderUserName, isSiteAdmin);
                Tenant.Context.ExecuteQuery();
            }
            catch (Exception e)
            {
                LogError(e, "Failed to set {0} site collection administrator permissions for site:{1}", userNameWithoutClaims, _siteUrl);
            }
        }
    }
}
