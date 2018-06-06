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
    class SkypeForBusinessPeerToPeerActivityUserCountsMap : ClassMap<SkypeForBusinessPeerToPeerActivityUserCounts>
    {
        SkypeForBusinessPeerToPeerActivityUserCountsMap()
        {
            Map(m => m.ReportRefreshDate).Name("Report Refresh Date").Index(0).Default(default(DateTime));
            Map(m => m.ReportDate).Name("Report Date").Index(1).Default(default(Nullable<DateTime>));
            Map(m => m.ReportPeriod).Name("Report Period").Index(2).Default(0);
            Map(m => m.IM).Name("IM").Index(3).Default(0);
            Map(m => m.Audio).Name("Audio").Index(4).Default(0);
            Map(m => m.Video).Name("Video").Index(5).Default(0);
            Map(m => m.AppSharing).Name("App Sharing").Index(6).Default(0);
            Map(m => m.FileTransfer).Name("File Transfer").Index(7).Default(0);
        }
    }

    /// <summary>
    /// Get usage trends on the number of unique users and type of peer-to-peer sessions held in your organization. Types of sessions include IM, audio, video, application sharing, and file transfers in peer-to-peer sessions.
    /// </summary>
    public class SkypeForBusinessPeerToPeerActivityUserCounts : JSONODataBase
    {
        [JsonProperty("reportRefreshDate")]
        public DateTime ReportRefreshDate { get; set; }

        [JsonProperty("reportDate")]
        public Nullable<DateTime> ReportDate { get; set; }

        [JsonProperty("reportPeriod")]
        public int ReportPeriod { get; set; }

        [JsonProperty("im")]
        public Nullable<Int64> IM { get; set; }

        [JsonProperty("audio")]
        public Nullable<Int64> Audio { get; set; }

        [JsonProperty("video")]
        public Nullable<Int64> Video { get; set; }

        [JsonProperty("appSharing")]
        public Nullable<Int64> AppSharing { get; set; }

        [JsonProperty("fileTransfer")]
        public Nullable<Int64> FileTransfer { get; set; }


    }
}
