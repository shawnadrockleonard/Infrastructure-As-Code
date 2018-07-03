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
Site Type,
Page View Count,
Report Date,
Report Period
2017-10-29,All,1335,2017-10-29,7
 */
    internal class SharePointSiteUsagePagesMap : ClassMap<SharePointSiteUsagePages>
    {
        internal SharePointSiteUsagePagesMap()
        {
            Map(m => m.ReportRefreshDate).Name("Report Refresh Date").Index(0).Default(default(DateTime));
            Map(m => m.SiteType).Name("Site Type").Index(1).Default(string.Empty);
            Map(m => m.PageViewCount).Name("Page View Count").Index(2).Default(string.Empty);
            Map(m => m.ReportDate).Name("Report Date").Index(3).Default(default(DateTime));
            Map(m => m.ReportPeriod).Name("Report Period").Index(4).Default(0);
        }
    }


    public class SharePointSiteUsagePages : JSONODataBase
    {
        [JsonProperty("reportRefreshDate")]
        public DateTime ReportRefreshDate { get; set; }

        [JsonProperty("siteType")]
        public string SiteType { get; set; }

        [JsonProperty("pageViewCount")]
        public Nullable<Int64> PageViewCount { get; set; }

        [JsonProperty("reportDate")]
        public DateTime ReportDate { get; set; }

        [JsonProperty("reportPeriod")]
        public int ReportPeriod { get; set; }
    }
}