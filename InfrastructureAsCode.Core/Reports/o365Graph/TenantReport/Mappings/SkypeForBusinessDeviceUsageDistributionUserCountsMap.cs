using CsvHelper.Configuration;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Reports.o365Graph.TenantReport.Mappings
{
    /*
     * CSV Mapping for Device usage
     * Report Refresh Date,,,,,,,,Report Period
     * */
    internal class SkypeForBusinessDeviceUsageDistributionUserCountsMap : ClassMap<SkypeForBusinessDeviceUsageDistributionUserCounts>
    {
        internal SkypeForBusinessDeviceUsageDistributionUserCountsMap()
        {
            Map(m => m.ReportRefreshDate).Name("Report Refresh Date").Index(0).Default(default(DateTime));
            Map(m => m.Windows).Name("Windows").Index(1).Default(0);
            Map(m => m.WindowsPhone).Name("Windows Phone").Index(2).Default(0);
            Map(m => m.AndroidPhone).Name("Android Phone").Index(3).Default(0);
            Map(m => m.iPhone).Name("iPhone").Index(4).Default(0);
            Map(m => m.iPad).Name("iPad").Index(5).Default(0);
            Map(m => m.ReportPeriod).Name("Report Period").Index(6).Default(0);
        }
    }

    /// <summary>
    /// Get the number of users using unique devices in your organization. The report will show you the number of users per device including Windows, Windows phone, Android phone, iPhone, and iPad.
    /// </summary>
    public class SkypeForBusinessDeviceUsageDistributionUserCounts : JSONODataBase
    {
        [JsonProperty("reportRefreshDate")]
        public DateTime ReportRefreshDate { get; set; }

        [JsonProperty("reportPeriod")]
        public int ReportPeriod { get; set; }

        [JsonProperty("windows")]
        public Nullable<Int64> Windows { get; set; }

        [JsonProperty("windowsPhone")]
        public Nullable<Int64> WindowsPhone { get; set; }

        [JsonProperty("androidPhone")]
        public Nullable<Int64> AndroidPhone { get; set; }

        [JsonProperty("iPhone")]
        public Nullable<Int64> iPhone { get; set; }

        [JsonProperty("iPad")]
        public Nullable<Int64> iPad { get; set; }
    }
}
