using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.Online.SharePoint.TenantManagement;

namespace InfrastructureAsCode.Core.Models
{
    /// <summary>
    /// Represents a Tenant Site Collection
    /// </summary>
    public class SPOSiteCollectionModel : CollectionModel
    {
        public SandboxedCodeActivationCapabilities Sandbox { get; set; }

        public string Title { get; set; }

        public double AverageResourceUsage { get; set; }

        public int CompatibilityLevel { get; set; }

        public double CurrentResourceUsage { get; set; }

        public CompanyWideSharingLinksPolicy DisableCompanyWideSharingLinks { get; set; }

        public string Owner { get; set; }

        public DateTime LastContentModifiedDate { get; set; }

        public DenyAddAndCustomizePagesStatus DenyAddAndCustomizePages { get; set; }

        public SharingCapabilities SharingCapability { get; set; }

        public string Status { get; set; }

        public long StorageMaximumLevel { get; set; }

        public long StorageUsage { get; set; }

        public int TimeZoneId { get; set; }

        public long StorageWarningLevel { get; set; }

        public string StorageQuotaType { get; set; }

        public double UserCodeMaximumLevel { get; set; }

        public double UserCodeWarningLevel { get; set; }

        public string Template { get; set; }
    }
}
