using InfrastructureAsCode.Core.Reports.o365Graph.TenantReport;
using CsvHelper.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace InfrastructureAsCode.Core.Reports.o365Graph.TenantReport.Mappings
{
    /*
     * 
The CSV file has the following headers for columns.
Report Refresh Date
Site URL
Owner Display Name
Is Deleted
Last Activity Date
File Count
Active File Count
Storage Used (Byte)
Storage Allocated (Byte)
Report Period
     */
    public class OneDriveUsageAccountDetailMap : ClassMap<OneDriveUsageAccountDetail>
    {
        public OneDriveUsageAccountDetailMap()
        {
            Map(m => m.ReportRefreshDate).Name("Report Refresh Date").Index(0).Default(default(DateTime));
            Map(m => m.SiteURL).Name("Site URL").Index(1).Default(string.Empty);
            Map(m => m.SiteOwner).Name("Owner Display Name").Index(2).Default(string.Empty);
            Map(m => m.Deleted).Name("Is Deleted").Index(3).Default(string.Empty);
            Map(m => m.LastActivityDateUTC).Name("Last Activity Date").Index(4).Default(default(DateTime));
            Map(m => m.Files).Name("File Count").Index(5).Default(0);
            Map(m => m.FilesViewedModified).Name("Active File Count").Index(6).Default(0);
            Map(m => m.StorageUsedInBytes).Name("Storage Used (Byte)").Index(7).Default(0);
            Map(m => m.StorageAllocatedInBytes).Name("Storage Allocated (Byte)").Index(8).Default(0);
            Map(m => m.ReportingPeriodDays).Name("Report Period").Index(9).Default(0);
        }
    }


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
        public Nullable<Int64> Files { get; set; }

        [JsonProperty("activeFileCount")]
        public Nullable<Int64> FilesViewedModified { get; set; }

        [JsonProperty("storageUsedInBytes")]
        public Nullable<Int64> StorageUsedInBytes { get; set; }

        [JsonProperty("storageAllocatedInBytes")]
        public Nullable<Int64> StorageAllocatedInBytes { get; set; }

        [JsonProperty("reportPeriod")]
        public int ReportingPeriodDays { get; set; }

    }
}
