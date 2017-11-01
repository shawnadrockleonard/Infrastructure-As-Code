using InfrastructureAsCode.Core.Models;
using InfrastructureAsCode.Core.Reports;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Extensions
{
    /// <summary>
    /// Provides extension methods to inject capability into a [Site]
    /// </summary>
    public static class SiteExtensions
    {
        /// <summary>
        /// Returns Taxonomy store and set details
        /// </summary>
        /// <param name="clientContext"></param>
        /// <param name="logger"></param>
        /// <param name="termSetId"></param>
        /// <returns>NULL if an Exception is thrown</returns>
        public static SPOTaxonomyModel GetTaxonomyFieldInfo(this ClientContext clientContext, ITraceLogger logger, Guid termSetId)
        {

            TaxonomySession session = TaxonomySession.GetTaxonomySession(clientContext);
            TermStore termStore = session.GetDefaultSiteCollectionTermStore();
            TermSet termSet = termStore.GetTermSet(termSetId);

            SPOTaxonomyTermStoreModel modelTermStore = null;
            SPOTaxonomyTermSetModel modelTermSet = null;
            try
            {
                clientContext.Load(termSet,
                    tctx => tctx.Id,
                    tctx => tctx.CustomSortOrder,
                    tctx => tctx.IsAvailableForTagging,
                    tctx => tctx.Owner,
                    tctx => tctx.CreatedDate,
                    tctx => tctx.LastModifiedDate,
                    tctx => tctx.Name,
                    tctx => tctx.Description,
                    tctx => tctx.TermStore,
                    tctx => tctx.Group,
                    tctx => tctx.IsOpenForTermCreation);

                clientContext.Load(termStore,
                    tctx => tctx.Id,
                    tctx => tctx.Name,
                    tctx => tctx.IsOnline,
                    tctx => tctx.DefaultLanguage,
                    tctx => tctx.ContentTypePublishingHub,
                    tctx => tctx.WorkingLanguage);
                clientContext.ExecuteQueryRetry();

                modelTermStore = new SPOTaxonomyTermStoreModel()
                {
                    Id = termStore.Id,
                    Name = termStore.Name,
                    IsOnline = termStore.IsOnline,
                    DefaultLanguage = termStore.DefaultLanguage,
                    ContentTypePublishingHub = termStore.ContentTypePublishingHub,
                    WorkingLanguage = termStore.WorkingLanguage
                };

                modelTermSet = new SPOTaxonomyTermSetModel()
                {
                    Id = termSet.Id,
                    IsAvailableForTagging = termSet.IsAvailableForTagging,
                    IsOpenForTermCreation = termSet.IsOpenForTermCreation,
                    CustomSortOrder = termSet.CustomSortOrder,
                    Owner = termSet.Owner,
                    CreatedDate = termSet.CreatedDate,
                    LastModifiedDate = termSet.LastModifiedDate,
                    Name = termSet.Name,
                    Description = termSet.Description
                };

                if(termSet.TermStore != null
                    && termSet.TermStore.Id != null)
                {
                    var tempStore = termSet.TermStore;
                    modelTermSet.TermStoreId = tempStore.Id;
                }

                if(termSet.Group != null)
                {
                    var termGroup = termSet.Group;
                    modelTermSet.Group = new SPOTaxonomyItemModel()
                    {
                        Id = termGroup.Id,
                        Name = termGroup.Name,
                        CreatedDate = termGroup.CreatedDate,
                        LastModifiedDate = termGroup.LastModifiedDate
                    };
                }
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Failed to retreive TermStore session {0} with message {1}", termSetId, ex.Message);
                return null;
            }

            // Build model
            var termsetModel = new SPOTaxonomyModel()
            {
                TermSetName = termSet.Name,
                TermSet = modelTermSet,
                TermStore = modelTermStore
            };

            return termsetModel;
        }
        
        /// <summary>
        /// Returns Taxonomy store and set details
        /// </summary>
        /// <param name="clientContext"></param>
        /// <param name="logger"></param>
        /// <param name="termSetName"></param>
        /// <param name="cultureId"></param>
        /// <returns>NULL if an Exception is thrown</returns>
        public static SPOTaxonomyModel GetTaxonomyFieldInfo(this ClientContext clientContext, ITraceLogger logger, string termSetName, int cultureId = 1033)
        {

            TaxonomySession session = TaxonomySession.GetTaxonomySession(clientContext);
            TermStore termStore = session.GetDefaultSiteCollectionTermStore();
            TermSetCollection termSets = termStore.GetTermSetsByName(termSetName, cultureId);

            SPOTaxonomyTermSetModel modelTermSet = null;
            SPOTaxonomyTermStoreModel modelTermStore = null;

            try
            {
                clientContext.Load(termSets, 
                    tsc => tsc.Include(
                        tctx => tctx.Id,
                        tctx => tctx.CustomSortOrder,
                        tctx => tctx.IsAvailableForTagging,
                        tctx => tctx.Owner,
                        tctx => tctx.CreatedDate,
                        tctx => tctx.LastModifiedDate,
                        tctx => tctx.Name,
                        tctx => tctx.Description,
                        tctx => tctx.TermStore,
                        tctx => tctx.Group,
                        tctx => tctx.IsOpenForTermCreation));

                clientContext.Load(termStore,
                    tctx => tctx.Id,
                    tctx => tctx.Name,
                    tctx => tctx.IsOnline,
                    tctx => tctx.DefaultLanguage,
                    tctx => tctx.ContentTypePublishingHub,
                    tctx => tctx.WorkingLanguage);
                clientContext.ExecuteQueryRetry();

                modelTermStore = new SPOTaxonomyTermStoreModel()
                {
                    Id = termStore.Id,
                    Name = termStore.Name,
                    IsOnline = termStore.IsOnline,
                    DefaultLanguage = termStore.DefaultLanguage,
                    ContentTypePublishingHub = termStore.ContentTypePublishingHub,
                    WorkingLanguage = termStore.WorkingLanguage
                };

                TermSet termSet = termSets.FirstOrDefault();
                modelTermSet = new SPOTaxonomyTermSetModel()
                {
                    Id = termSet.Id,
                    IsAvailableForTagging = termSet.IsAvailableForTagging,
                    IsOpenForTermCreation = termSet.IsOpenForTermCreation,
                    CustomSortOrder = termSet.CustomSortOrder,
                    Owner = termSet.Owner,
                    CreatedDate = termSet.CreatedDate,
                    LastModifiedDate = termSet.LastModifiedDate,
                    Name = termSet.Name,
                    Description = termSet.Description
                };

                if (termSet.TermStore != null
                    && termSet.TermStore.Id != null)
                {
                    var tempStore = termSet.TermStore;
                    modelTermSet.TermStoreId = tempStore.Id;
                }

                if (termSet.Group != null)
                {
                    var termGroup = termSet.Group;
                    modelTermSet.Group = new SPOTaxonomyItemModel()
                    {
                        Id = termGroup.Id,
                        Name = termGroup.Name,
                        CreatedDate = termGroup.CreatedDate,
                        LastModifiedDate = termGroup.LastModifiedDate
                    };
                }
            }
            catch(Exception ex)
            {
                logger.LogError(ex, "Failed to retreive TermStore session {0} with message {1}", termSetName, ex.Message);
                return null;
            }

            // Build model
            var termsetModel = new SPOTaxonomyModel()
            {
                TermSetName = termSetName,
                TermStore = modelTermStore,
                TermSet = modelTermSet
            };

            return termsetModel;
        }

        /// <summary>
        /// Adds or Updates an existing Custom Action [ScriptSrc] into the [Site] Custom Actions
        /// </summary>
        /// <param name="site"></param>
        /// <param name="customactionname"></param>
        /// <param name="customactionurl"></param>
        /// <param name="sequence"></param>
        public static void AddOrUpdateCustomActionLink(this Site site, string customactionname, string customactionurl, int sequence)
        {
            var sitecustomActions = site.GetCustomActions();
            UserCustomAction cssAction = null;
            if (site.CustomActionExists(customactionname))
            {
                cssAction = sitecustomActions.FirstOrDefault(fod => fod.Name == customactionname);
            }
            else
            {
                // Build a custom action to write a link to our new CSS file
                cssAction = site.UserCustomActions.Add();
                cssAction.Name = customactionname;
                cssAction.Location = "ScriptLink";
            }

            cssAction.Sequence = sequence;
            cssAction.ScriptSrc = customactionurl;
            cssAction.Update();
            site.Context.ExecuteQueryRetry();
        }

        /// <summary>
        /// Adds or Updates an existing Custom Action [ScriptBlock] into the [Site] Custom Actions
        /// </summary>
        /// <param name="site"></param>
        /// <param name="customactionname"></param>
        /// <param name="customActionBlock"></param>
        /// <param name="sequence"></param>
        public static void AddOrUpdateCustomActionLinkBlock(this Site site, string customactionname, string customActionBlock, int sequence)
        {
            var sitecustomActions = site.GetCustomActions();
            UserCustomAction cssAction = null;
            if (site.CustomActionExists(customactionname))
            {
                cssAction = sitecustomActions.FirstOrDefault(fod => fod.Name == customactionname);
            }
            else
            {
                // Build a custom action to write a link to our new CSS file
                cssAction = site.UserCustomActions.Add();
                cssAction.Name = customactionname;
                cssAction.Location = "ScriptLink";
            }

            cssAction.Sequence = sequence;
            cssAction.ScriptBlock = customActionBlock;
            cssAction.Update();
            site.Context.ExecuteQueryRetry();
        }

        /// <summary>
        /// Will remove the custom action if one exists
        /// </summary>
        /// <param name="site"></param>
        /// <param name="customactionname"></param>
        public static bool RemoveCustomActionLink(this Site site, string customactionname)
        {
            if (site.CustomActionExists(customactionname))
            {
                var cssAction = site.GetCustomActions().FirstOrDefault(fod => fod.Name == customactionname || fod.Title == customactionname);
                site.DeleteCustomAction(cssAction.Id);
                return true;
            }
            return false;
        }

        /// <summary>
        /// Extracts the Site Usage
        /// </summary>
        /// <remarks>
        /// If the Usage property has not be instantiated it will load the property into memory; 
        ///     Note: This will slow the processing
        /// </remarks>
        /// <param name="site"></param>
        /// <returns></returns>
        public static SPSiteUsageModel GetSiteUsageMetric(this Site site)
        {

            if (!site.IsPropertyAvailable(sctx => sctx.Usage))
            {
                site.EnsureProperty(sctx => sctx.Usage);
            }

            var byteAddressSpace = 1024;
            var mbAddressSpace = 1048576;

            UsageInfo _usageInfo = site.Usage;
            Double _storageBytes = _usageInfo.Storage;
            var _storage = (_storageBytes / mbAddressSpace);
            Double _storageQuotaBytes = (_usageInfo.StoragePercentageUsed > 0) ? _storageBytes / _usageInfo.StoragePercentageUsed : 0;
            var _storageQuota = (_storageQuotaBytes / mbAddressSpace);

            var _storageUsageMB = Math.Round(_storage, 4);
            var _storageAllocatedMB = Math.Round(_storageQuota, 4);

            var _storageUsageGB = Math.Round((_storage / byteAddressSpace), 4);
            var _storageAllocatedGB = Math.Round((_storageQuota / byteAddressSpace), 4);

            var _storagePercentUsed = Math.Round(_usageInfo.StoragePercentageUsed, 4);


            var storageModel = new SPSiteUsageModel()
            {
                Bandwidth = _usageInfo.Bandwidth,
                DiscussionStorage = _usageInfo.DiscussionStorage,
                Hits = _usageInfo.Hits,
                Visits = _usageInfo.Visits,
                StorageQuotaBytes = _storageQuotaBytes,
                AllocatedGb = _storageAllocatedGB,
                AllocatedGbDecimal = _storageAllocatedGB.TryParseDecimal(0),
                UsageGb = _storageUsageGB,
                UsageGbDecimal = _storageUsageGB.TryParseDecimal(0),
                AllocatedMb = _storageAllocatedMB,
                AllocatedMbDecimal = _storageAllocatedMB.TryParseDecimal(0),
                UsageMb = _storageUsageMB,
                UsageMbDecimal = _storageUsageMB.TryParseDecimal(0),
                StorageUsedPercentage = _storagePercentUsed,
                StorageUsedPercentageDecimal = _storagePercentUsed.TryParseDecimal(0)
            };

            return storageModel;
        }
    }
}
