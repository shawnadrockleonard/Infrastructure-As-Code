using InfrastructureAsCode.Core.Models;
using Microsoft.Online.SharePoint.TenantAdministration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;

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
    }
}
