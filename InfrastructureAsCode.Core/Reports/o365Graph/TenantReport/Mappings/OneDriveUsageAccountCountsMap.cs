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
Site Type
Total
Active
Report Date
Report Period
     */
    internal class OneDriveUsageAccountCountsMap : ClassMap<OneDriveUsageAccountCounts>
    {
        internal OneDriveUsageAccountCountsMap()
        {
            Map(m => m.ReportRefreshDate).Name("Report Refresh Date").Index(0).Default(default(DateTime));
            Map(m => m.SiteType).Name("Site Type").Index(1).Default(string.Empty);
            Map(m => m.Total_Accounts).Name("Total").Index(2).Default(0);
            Map(m => m.Active_Accounts).Name("Active").Index(3).Default(0);
            Map(m => m.ReportDate).Name("Report Date").Index(4).Default(default(DateTime));
            Map(m => m.ReportPeriod).Name("Report Period").Index(5).Default(0);
        }
    }


    public class OneDriveUsageAccountCounts : JSONODataBase
    {
        [JsonProperty("reportRefreshDate")]
        public DateTime ReportRefreshDate { get; set; }

        [JsonProperty("reportDate")]
        public DateTime ReportDate { get; set; }

        [JsonProperty("siteType")]
        public string SiteType { get; set; }

        [JsonProperty("total")]
        public Nullable<Int64> Total_Accounts { get; set; }

        [JsonProperty("active")]
        public Nullable<Int64> Active_Accounts { get; set; }

        [JsonProperty("reportPeriod")]
        public int ReportPeriod { get; set; }

    }
}
