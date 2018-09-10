using InfrastructureAsCode.Core.Enums;
using InfrastructureAsCode.Core.Models;
using InfrastructureAsCode.Core.Extensions;
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.Online.SharePoint.TenantManagement;
using Microsoft.PowerShell.Commands;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using Resources = InfrastructureAsCode.Core.Properties.Resources;

namespace InfrastructureAsCode.Powershell.Commands.Base
{
    /// <summary>
    /// SharePoint Online base command for tenant level administration
    /// </summary>
    public abstract class IaCAdminCmdlet : IaCCmdlet
    {
        private Microsoft.Online.SharePoint.TenantAdministration.Tenant _tenant;
        /// <summary>
        /// Create client context to tenant admin
        /// </summary>
        public Microsoft.Online.SharePoint.TenantAdministration.Tenant TenantContext
        {
            get
            {
                if (_tenant == null)
                {
                    _tenant = new Microsoft.Online.SharePoint.TenantAdministration.Tenant(this.ClientContext);
                }
                return _tenant;
            }
        }

        private Office365Tenant _officeTenant { get; set; }

        /// <summary>
        /// Initializes a Office365 Tenant Context
        /// </summary>
        public Office365Tenant OfficeTenantContext
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

        /// <summary>
        /// Initializers the logger from the cmdlet
        /// </summary>
        protected override void OnBeginInitialize()
        {
            base.OnBeginInitialize();

            SPIaCConnection.CurrentConnection.CacheContext();

            Uri uri = new Uri(this.ClientContext.Url);
            var urlParts = uri.Authority.Split(new[] { '.' });
            if (!urlParts[0].EndsWith("-admin"))
            {
                var adminUrl = string.Format("https://{0}-admin.{1}.{2}", urlParts[0], urlParts[1], urlParts[2]);

                SPIaCConnection.CurrentConnection.Context = this.ClientContext.Clone(adminUrl);
            }


            if (TenantContext.Context == null)
            {
                this.ClientContext.Load(TenantContext);
                this.ClientContext.ExecuteQueryRetry();
            }
        }

        protected override void EndProcessing()
        {
            SPIaCConnection.CurrentConnection.RestoreCachedContext();
            if (TenantContext.Context == null)
            {
                TenantContext.Context.Dispose();
            }

            base.EndProcessing();
        }

        /// <summary>
        /// Sets the site collection administrator for the activity
        /// </summary>
        /// <param name="_siteUrl">The relative url to the site collection</param>
        /// <param name="userNameWithoutClaims">Provide the username without the claim prefix</param>
        /// <param name="isSiteAdmin">(OPTIONAL) true to set the user as a site collection administrator</param>
        protected virtual void SetSiteAdmin(string _siteUrl, string userNameWithoutClaims, bool isSiteAdmin = false)
        {
            var claimProviderUserName = string.Format("{0}|{1}", ClaimIdentifier, userNameWithoutClaims);
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
                if (TenantContext.Context == null)
                {
                    this.ClientContext.Load(TenantContext);
                    this.ClientContext.ExecuteQueryRetry();
                }
                TenantContext.SetSiteAdmin(_siteUrl, claimProviderUserName, isSiteAdmin);
                TenantContext.Context.ExecuteQueryRetry();
            }
            catch (Exception e)
            {
                LogError(e, "Failed to set {0} site collection administrator permissions for site:{1}", userNameWithoutClaims, _siteUrl);
            }
        }

        /// <summary>
        /// Returns all site collections in the tenant
        /// </summary>
        /// <param name="includeProperties">Include all Site Collection properties</param>
        /// <returns></returns>
        public List<SPOSiteCollectionModel> GetSiteCollections(bool includeProperties = false)
        {
            var urls = TenantContext.GetSPOSiteCollections(includeProperties);
            LogVerbose("Found URLs {0}", urls.Count);
            return urls;
        }

        /// <summary>
        /// removes claim prefix from the user logon
        /// </summary>
        /// <param name="_user"></param>
        /// <returns></returns>
        internal string RemoveClaimIdentifier(string _user)
        {
            var _cleanedUser = _user;

            var _tmp = _user.Split(new char[] { '|' });
            if (_tmp.Length > 0)
            {
                _cleanedUser = _tmp.Last(); // remove claim identity
            }

            return _cleanedUser;
        }
    }
}
