using InfrastructureAsCode.Core.Models;
using Microsoft.Online.SharePoint.TenantAdministration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using Microsoft.Online.SharePoint.TenantManagement;
using InfrastructureAsCode.Core.Reports;

namespace InfrastructureAsCode.Core.Extensions
{
    /// <summary>
    /// Provides global extension functionality
    /// </summary>
    public static class TenantExtensions
    {
        /// <summary>
        /// Returns all site collections in the tenant
        /// </summary>
        /// <param name="includeProperties">Include all Site Collection properties</param>
        /// <returns></returns>
        public static List<SPOSiteCollectionModel> GetSPOSiteCollections(this Tenant tenantContext, bool includeProperties = false)
        {
            int startIndex = 0; var urls = new List<SPOSiteCollectionModel>();
            SPOSitePropertiesEnumerable spenumerable = null;
            while (spenumerable == null || spenumerable.Count > 0)
            {
                spenumerable = tenantContext.GetSiteProperties(startIndex, includeProperties);
                tenantContext.Context.Load(spenumerable);
                tenantContext.Context.ExecuteQueryRetry();

                foreach (SiteProperties sp in spenumerable)
                {
                    SPOSiteCollectionModel model = null;
                    if (includeProperties)
                    {
                        model = new SPOSiteCollectionModel()
                        {
                            Url = sp.Url.EnsureTrailingSlashLowered(),
                            title = sp.Title,
                            sandbox = sp.SandboxedCodeActivationCapability,
                            AverageResourceUsage = sp.AverageResourceUsage,
                            CompatibilityLevel = sp.CompatibilityLevel,
                            CurrentResourceUsage = sp.CurrentResourceUsage,
                            DenyAddAndCustomizePages = sp.DenyAddAndCustomizePages,
                            DisableCompanyWideSharingLinks = sp.DisableCompanyWideSharingLinks,
                            LastContentModifiedDate = sp.LastContentModifiedDate,
                            Owner = sp.Owner,
                            SharingCapability = sp.SharingCapability,
                            Status = sp.Status,
                            StorageMaximumLevel = sp.StorageMaximumLevel,
                            StorageQuotaType = sp.StorageQuotaType,
                            StorageUsage = sp.StorageUsage,
                            StorageWarningLevel = sp.StorageWarningLevel,
                            TimeZoneId = sp.TimeZoneId,
                            WebsCount = sp.WebsCount,
                            Template = sp.Template,
                            UserCodeWarningLevel = sp.UserCodeWarningLevel,
                            UserCodeMaximumLevel = sp.UserCodeMaximumLevel
                        };
                    }
                    else
                    {
                        model = new SPOSiteCollectionModel()
                        {
                            Url = sp.Url.EnsureTrailingSlashLowered(),
                            title = sp.Title
                        };
                    }

                    urls.Add(model);
                }
                startIndex += spenumerable.Count;
            }

            return urls;
        }

        /// <summary>
        /// Query the Tenant UPS based on Site Collection
        /// </summary>
        /// <param name="siteUrl"></param>
        /// <param name="invitedAs"></param>
        /// <returns></returns>
        public static List<SPExternalUserEntity> CheckExternalUserForSite(this ClientContext adminContext, ITraceLogger logger, string siteUrl, string invitedAs = "")
        {
            if (siteUrl == null)
                throw new ArgumentNullException("siteUrl");


            var externalUsers = new List<SPExternalUserEntity>();
            int pageSize = 50;
            int position = 0;
            GetExternalUsersResults results = null;

            var officeTenantContext = new Office365Tenant(adminContext);

            while (true)
            {
                logger.LogInformation($"Checking External User with {invitedAs} at start {position} and page size {pageSize}");

                results = officeTenantContext.GetExternalUsersForSite(siteUrl, position, pageSize, invitedAs, SortOrder.Ascending);
                adminContext.Load(results, r => r.UserCollectionPosition, r => r.TotalUserCount, r => r.ExternalUserCollection);
                adminContext.ExecuteQueryRetry();

                foreach (ExternalUser externalUser in results.ExternalUserCollection)
                {
                    externalUsers.Add(new SPExternalUserEntity()
                    {
                        AcceptedAs = externalUser.AcceptedAs,
                        DisplayName = externalUser.DisplayName,
                        InvitedAs = externalUser.InvitedAs,
                        InvitedBy = externalUser.InvitedBy,
                        UniqueId = externalUser.UniqueId,
                        UserId = externalUser.UserId,
                        WhenCreated = externalUser.WhenCreated
                    });
                }

                position = results.UserCollectionPosition;

                if (position == -1 || position == results.TotalUserCount)
                {
                    break;
                }
            }

            return externalUsers;
        }

        /// <summary>
        /// Query the Tenant UPS
        /// </summary>
        /// <param name="invitedAs"></param>
        /// <returns></returns>
        public static List<SPExternalUserEntity> CheckExternalUser(this ClientContext adminContext, ITraceLogger logger, string invitedAs = "")
        {
            var externalUsers = new List<SPExternalUserEntity>();
            int pageSize = 50;
            int position = 0;
            GetExternalUsersResults results = null;

            var officeTenantContext = new Office365Tenant(adminContext);


            while (true)
            {
                logger.LogInformation($"Checking External User with {invitedAs} at start {position} and page size {pageSize}");

                results = officeTenantContext.GetExternalUsers(position, pageSize, invitedAs, Microsoft.Online.SharePoint.TenantManagement.SortOrder.Ascending);
                adminContext.Load(results, r => r.UserCollectionPosition, r => r.TotalUserCount, r => r.ExternalUserCollection);
                adminContext.ExecuteQueryRetry();

                foreach (ExternalUser externalUser in results.ExternalUserCollection)
                {
                    externalUsers.Add(new SPExternalUserEntity()
                    {
                        DisplayName = externalUser.DisplayName,
                        AcceptedAs = externalUser.AcceptedAs,
                        InvitedAs = externalUser.InvitedAs,
                        InvitedBy = externalUser.InvitedBy,
                        UniqueId = externalUser.UniqueId,
                        UserId = externalUser.UserId,
                        WhenCreated = externalUser.WhenCreated
                    });
                }

                position = results.UserCollectionPosition;

                if (position == -1 || position == results.TotalUserCount)
                {
                    break;
                }
            }

            return externalUsers;
        }
    }
}
