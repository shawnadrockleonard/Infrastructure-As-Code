using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Models
{
    public class ExtendSPOSiteModel : SPOSiteCollectionModel
    {
        public ExtendSPOSiteModel() : base()
        {
            this.Sites = new List<SPSiteModel>();
        }

        public ExtendSPOSiteModel(SPOSiteCollectionModel self) : this()
        {
            this.AverageResourceUsage = self.AverageResourceUsage;
            this.Url = self.Url;
            this.title = self.title;
            this.sandbox = self.sandbox;
            this.AverageResourceUsage = self.AverageResourceUsage;
            this.CompatibilityLevel = self.CompatibilityLevel;
            this.CurrentResourceUsage = self.CurrentResourceUsage;
            this.DenyAddAndCustomizePages = self.DenyAddAndCustomizePages;
            this.DisableCompanyWideSharingLinks = self.DisableCompanyWideSharingLinks;
            this.LastContentModifiedDate = self.LastContentModifiedDate;
            this.Owner = self.Owner;
            this.SharingCapability = self.SharingCapability;
            this.Status = self.Status;
            this.StorageMaximumLevel = self.StorageMaximumLevel;
            this.StorageQuotaType = self.StorageQuotaType;
            this.StorageUsage = self.StorageUsage;
            this.StorageWarningLevel = self.StorageWarningLevel;
            this.TimeZoneId = self.TimeZoneId;
            this.WebsCount = self.WebsCount;
            this.Template = self.Template;
            this.UserCodeWarningLevel = self.UserCodeWarningLevel;
            this.UserCodeMaximumLevel = self.UserCodeMaximumLevel;
        }

        public IList<SPSiteModel> Sites { get; set; }


    }
}
