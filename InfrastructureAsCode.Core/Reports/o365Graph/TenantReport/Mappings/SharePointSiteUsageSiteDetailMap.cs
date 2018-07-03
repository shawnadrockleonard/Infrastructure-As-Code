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
Report Refresh Date,
Site URL,
Owner Display Name,
Is Deleted,
Last Activity Date,
File Count,
Active File Count,
Page View Count,
Visited Page Count,
Storage Used (Byte),
Storage Allocated (Byte),
Root Web Template,
Report Period
     */
    internal class SharePointSiteUsageSiteDetailMap : ClassMap<SharePointSiteUsageSiteDetail>
    {
        internal SharePointSiteUsageSiteDetailMap()
        {
            Map(m => m.ReportRefreshDate).Name("Report Refresh Date").Index(0).Default(default(DateTime));
            Map(m => m.SiteURL).Name("Site URL").Index(1).Default(string.Empty);
            Map(m => m.OwnerDisplayName).Name("Owner Display Name").Index(2).Default(string.Empty);
            Map(m => m.IsDeleted).Name("Is Deleted").Index(3).Default(false);
            Map(m => m.LastActivityDate).Name("Last Activity Date").Index(4).Default(default(Nullable<DateTime>));
            Map(m => m.FileCount).Name("File Count").Index(5).Default(0);
            Map(m => m.ActiveFileCount).Name("Active File Count").Index(6).Default(0);
            Map(m => m.PageViewCount).Name("Page View Count").Index(7).Default(0);
            Map(m => m.VisitedPageCount).Name("Visited Page Count").Index(8).Default(0);
            Map(m => m.StorageUsed_Byte).Name("Storage Used (Byte)").Index(9).Default(0);
            Map(m => m.StorageAllocated_Byte).Name("Storage Allocated (Byte)").Index(10).Default(0);
            Map(m => m.RootWebTemplate).Name("Root Web Template").Index(11).Default(string.Empty);
            Map(m => m.ReportPeriod).Name("Report Period").Index(12).Default(0);
        }
    }



    public class SharePointSiteUsageSiteDetail : JSONODataBase
    {
        [JsonProperty("reportRefreshDate")]
        public DateTime ReportRefreshDate { get; set; }

        [JsonProperty("siteId")]
        public string SiteID { get; set; }

        [JsonProperty("siteUrl")]
        public string SiteURL { get; set; }

        [JsonProperty("ownerDisplayName")]
        public string OwnerDisplayName { get; set; }

        [JsonProperty("isDeleted")]
        public bool IsDeleted { get; set; }

        [JsonProperty("lastActivityDate")]
        public Nullable<DateTime> LastActivityDate { get; set; }

        [JsonProperty("fileCount")]
        public Nullable<Int64> FileCount { get; set; }

        [JsonProperty("activeFileCount")]
        public Nullable<Int64> ActiveFileCount { get; set; }

        [JsonProperty("pageViewCount")]
        public Nullable<Int64> PageViewCount { get; set; }

        [JsonProperty("visitedPageCount")]
        public Nullable<Int64> VisitedPageCount { get; set; }

        [JsonProperty("storageUsedInBytes")]
        public Nullable<Int64> StorageUsed_Byte { get; set; }

        [JsonProperty("storageAllocatedInBytes")]
        public Nullable<Int64> StorageAllocated_Byte { get; set; }

        [JsonProperty("rootWebTemplate")]
        public string RootWebTemplate { get; set; }

        [JsonProperty("reportPeriod")]
        public int ReportPeriod { get; set; }
    }
}
