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
using InfrastructureAsCode.Core.HttpServices;

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
                            Title = sp.Title,
                            Sandbox = sp.SandboxedCodeActivationCapability,
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
                            Title = sp.Title
                        };
                    }

                    urls.Add(model);
                }
                startIndex += spenumerable.Count;
            }

            return urls;
        }

        /// <summary>
        /// Retreive all of the OneDrive for Business profiles
        /// </summary>
        /// <param name="adminSiteContext"></param>
        /// <param name="traceLogger"></param>
        /// <param name="MySiteUrl"></param>
        /// <param name="includeProperties"></param>
        /// <returns></returns>
        public static List<OD4BSiteCollectionModel> GetOneDriveSiteCollections(this ClientContext adminSiteContext, ITraceLogger traceLogger, string MySiteUrl, bool includeProperties = false)
        {
            var results = new List<OD4BSiteCollectionModel>();
            MySiteUrl = MySiteUrl.EnsureTrailingSlashLowered();
            MySiteUrl = MySiteUrl.Substring(0, MySiteUrl.Length - 1);

            using (var _UserProfileService = new UserProfileService(adminSiteContext, adminSiteContext.Url))
            {
                var userProfileResult = _UserProfileService.OWService.GetUserProfileByIndex(-1);
                var userProfilesCount = _UserProfileService.OWService.GetUserProfileCount();
                var rowIndex = 1;

                // As long as the next User profile is NOT the one we started with (at -1)...
                while (int.TryParse(userProfileResult.NextValue, out int nextValueIndex) && nextValueIndex != -1)
                {
                    if ((rowIndex % 50) == 0 || rowIndex > userProfilesCount)
                    {
                        traceLogger.LogInformation($"Next set {rowIndex} of {userProfilesCount}");
                    }

                    try
                    {
                        var personalSpace = userProfileResult.RetrieveUserProperty("PersonalSpace");
                        var personalSpaceUrl = $"{MySiteUrl}{personalSpace}";
                        var model = new OD4BSiteCollectionModel();

                        if (includeProperties == false)
                        {
                            model = new OD4BSiteCollectionModel
                            {
                                PersonalSpaceProperty = personalSpace,
                                Url = personalSpaceUrl,
                                NameProperty = userProfileResult.RetrieveUserProperty("PreferredName"),
                                UserName = userProfileResult.RetrieveUserProperty("UserName"),
                                Title = userProfileResult.RetrieveUserProperty("Title")
                            };
                        }
                        else
                        {
                            model = new OD4BSiteCollectionModel
                            {
                                PersonalSpaceProperty = personalSpace,
                                Url = personalSpaceUrl,
                                NameProperty = userProfileResult.RetrieveUserProperty("PreferredName"),
                                UserName = userProfileResult.RetrieveUserProperty("UserName"),
                                PictureUrl = userProfileResult.RetrieveUserProperty("PictureURL"),
                                AboutMe = userProfileResult.RetrieveUserProperty("AboutMe"),
                                SpsSkills = userProfileResult.RetrieveUserProperty("SPS-Skills"),
                                Manager = userProfileResult.RetrieveUserProperty("Manager"),
                                WorkPhone = userProfileResult.RetrieveUserProperty("WorkPhone"),
                                Department = userProfileResult.RetrieveUserProperty("Department"),
                                Company = userProfileResult.RetrieveUserProperty("Company"),
                                AccountName = userProfileResult.RetrieveUserProperty("AccountName"),
                                DistinguishedName = userProfileResult.RetrieveUserProperty("SPS-DistinguishedName"),
                                FirstName = userProfileResult.RetrieveUserProperty("FirstName"),
                                LastName = userProfileResult.RetrieveUserProperty("LastName"),
                                UserPrincipalName = userProfileResult.RetrieveUserProperty("SPS-UserPrincipalName"),
                                Title = userProfileResult.RetrieveUserProperty("Title"),
                                WorkEmail = userProfileResult.RetrieveUserProperty("WorkEmail"),
                                HomePhone = userProfileResult.RetrieveUserProperty("HomePhone"),
                                CellPhone = userProfileResult.RetrieveUserProperty("CellPhone"),
                                Office = userProfileResult.RetrieveUserProperty("Office"),
                                Location = userProfileResult.RetrieveUserProperty("SPS-Location"),
                                Fax = userProfileResult.RetrieveUserProperty("Fax"),
                                MailingAddress = userProfileResult.RetrieveUserProperty("MailingAddress"),
                                School = userProfileResult.RetrieveUserProperty("SPS-School"),
                                WebSite = userProfileResult.RetrieveUserProperty("WebSite"),
                                Education = userProfileResult.RetrieveUserProperty("Education"),
                                JobTitle = userProfileResult.RetrieveUserProperty("SPS-JobTitle"),
                                Assistant = userProfileResult.RetrieveUserProperty("Assistant"),
                                HireDate = userProfileResult.RetrieveUserProperty("SPS-HireDate"),
                                TimeZone = userProfileResult.RetrieveUserProperty("SPS-TimeZone"),
                                Locale = userProfileResult.RetrieveUserProperty("SPS-Locale"),
                                EmailOptin = userProfileResult.RetrieveUserProperty("SPS-EmailOptin"),
                                PrivacyPeople = userProfileResult.RetrieveUserProperty("SPS-PrivacyPeople"),
                                PrivacyActivity = userProfileResult.RetrieveUserProperty("SPS-PrivacyActivity"),
                                MySiteUpgrade = userProfileResult.RetrieveUserProperty("SPS-MySiteUpgrade"),
                                ProxyAddresses = userProfileResult.RetrieveUserProperty("SPS-ProxyAddresses"),
                                OWAUrl = userProfileResult.RetrieveUserProperty("SPS-OWAUrl")
                            };
                        }
                        results.Add(model);

                        userProfileResult = _UserProfileService.OWService.GetUserProfileByIndex(int.Parse(userProfileResult.NextValue));
                        rowIndex++;
                    }
                    catch (Exception e)
                    {
                        traceLogger.LogWarning("Failed to execute while loop {0}", e.Message);
                    }
                }

                // Final processing
                traceLogger.LogWarning($"Total Profiles {rowIndex} processed...");
            }

            return results;
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
