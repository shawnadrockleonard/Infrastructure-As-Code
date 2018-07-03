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
     * CSV Mapping
     * Report Refresh Date,Report Date,Report Period,,,,,
     * */
    internal class SkypeForBusinessPeerToPeerActivityMinuteCountsMap : ClassMap<SkypeForBusinessPeerToPeerActivityMinuteCounts>
    {
        internal SkypeForBusinessPeerToPeerActivityMinuteCountsMap()
        {
            Map(m => m.ReportRefreshDate).Name("Report Refresh Date").Index(0).Default(default(DateTime));
            Map(m => m.ReportDate).Name("Report Date").Index(1).Default(default(Nullable<DateTime>));
            Map(m => m.ReportPeriod).Name("Report Period").Index(2).Default(0);
            Map(m => m.Audio).Name("Audio").Index(3).Default(0);
            Map(m => m.Video).Name("Video").Index(4).Default(0);
        }
    }

    /// <summary>
    /// Get usage trends on the length in minutes and type of peer-to-peer sessions held in your organization. Types of sessions include audio and video.
    /// </summary>
    public class SkypeForBusinessPeerToPeerActivityMinuteCounts : JSONODataBase
    {
        [JsonProperty("reportRefreshDate")]
        public DateTime ReportRefreshDate { get; set; }

        [JsonProperty("reportDate")]
        public Nullable<DateTime> ReportDate { get; set; }

        [JsonProperty("reportPeriod")]
        public int ReportPeriod { get; set; }

        [JsonProperty("audio")]
        public Nullable<Int64> Audio { get; set; }

        [JsonProperty("video")]
        public Nullable<Int64> Video { get; set; }
    }
}
