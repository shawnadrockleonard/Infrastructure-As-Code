using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Reports.o365Graph.TenantReport
{
    public class OneDriveUsageAccountDetail : JSONODataBase
    {
        [JsonProperty("reportRefreshDate")]
        public DateTime ReportRefreshDate { get; set; }

        [JsonProperty("lastActivityDate")]
        public DateTime LastActivityDateUTC { get; set; }

        [JsonProperty("siteUrl")]
        public string SiteURL { get; set; }

        [JsonProperty("ownerDisplayName")]
        public string SiteOwner { get; set; }

        [JsonProperty("isDeleted")]
        public string Deleted { get; set; }

        [JsonProperty("fileCount")]
        public Int64 Files { get; set; }

        [JsonProperty("activeFileCount")]
        public Int64 FilesViewedModified { get; set; }

        [JsonProperty("storageUsedInBytes")]
        public Int64 StorageUsedInBytes { get; set; }

        [JsonProperty("storageAllocatedInBytes")]
        public Int64 StorageAllocatedInBytes { get; set; }

        [JsonProperty("reportPeriod")]
        public int ReportingPeriodDays { get; set; }

    }
}
