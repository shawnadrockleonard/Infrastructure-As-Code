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
The CSV file has the following headers for columns.
Report Refresh Date
Site Type
Storage Used (Byte)
Report Date
Report Period
     */
    internal class OneDriveUsageStorageMap : ClassMap<OneDriveUsageStorage>
    {
        internal OneDriveUsageStorageMap()
        {
            Map(m => m.ReportRefreshDate).Name("Report Refresh Date").Index(0).Default(default(DateTime));
            Map(m => m.SiteType).Name("Site Type").Index(1).Default(string.Empty);
            Map(m => m.StorageUsed_Bytes).Name("Storage Used (Byte)").Index(2).Default(0);
            Map(m => m.ReportDate).Name("Report Date").Index(3).Default(default(DateTime));
            Map(m => m.ReportingPeriodDays).Name("Report Period").Index(4).Default(0);
        }
    }



    public class OneDriveUsageStorage : JSONODataBase
    {
        [JsonProperty("reportRefreshDate")]
        public DateTime ReportRefreshDate { get; set; }

        [JsonProperty("reportDate")]
        public DateTime ReportDate { get; set; }

        [JsonProperty("siteType")]
        public string SiteType { get; set; }

        [JsonProperty("storageUsedInBytes")]
        public Nullable<Int64> StorageUsed_Bytes { get; set; }

        [JsonProperty("reportPeriod")]
        public int ReportingPeriodDays { get; set; }

    }
}
