using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Models
{
    /// <summary>
    /// Represents a breakdown of the Site Quota and Usage metrics
    /// </summary>
    public class SPSiteUsageModel
    {
        /// <summary>
        /// Bandwidth usage
        /// </summary>
        public long Bandwidth { get; set; }

        /// <summary>
        /// Total storage occupied by Discussions
        /// </summary>
        public long DiscussionStorage { get; set; }

        /// <summary>
        /// Total hits to the site
        /// </summary>
        public long Hits { get; set; }

        /// <summary>
        /// Total visits to the site
        /// </summary>
        public long Visits { get; set; }

        /// <summary>
        /// Represents a calculated representation of the quota based on Storage Percentage and Storage Used
        /// </summary>
        public double StorageQuotaBytes { get; set; }

        /// <summary>
        /// Calculated MB based on Quota/Percentage Used
        /// </summary>
        public double AllocatedMb { get; set; }

        /// <summary>
        /// Rounded and parsed to decimal
        /// </summary>
        public decimal AllocatedMbDecimal { get; set; }

        /// <summary>
        /// Total Storage used in MB
        /// </summary>
        public double UsageMb { get; set; }

        /// <summary>
        /// Rounded and parsed to decimal
        /// </summary>
        public decimal UsageMbDecimal { get; set; }

        /// <summary>
        /// Calculated GB based on Quota/Percentage Used
        /// </summary>
        public double AllocatedGb { get; set; }

        /// <summary>
        /// Rounded and parsed to decimal
        /// </summary>
        public decimal AllocatedGbDecimal { get; set; }

        /// <summary>
        /// Total Storage used in GB
        /// </summary>
        public double UsageGb { get; set; }

        /// <summary>
        /// Rounded and parsed to decimal
        /// </summary>
        public decimal UsageGbDecimal { get; set; }

        /// <summary>
        /// Storage percentage pulled from Site Usage object
        /// </summary>
        public double StorageUsedPercentage { get; set; }

        /// <summary>
        /// Rounded and parsed to decimal
        /// </summary>
        public decimal StorageUsedPercentageDecimal { get; set; }

    }
}
