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
    internal class SkypeForBusinessParticipantActivityUserCountsMap : ClassMap<SkypeForBusinessParticipantActivityUserCounts>
    {
        internal SkypeForBusinessParticipantActivityUserCountsMap()
        {
            Map(m => m.ReportRefreshDate).Name("Report Refresh Date").Index(0).Default(default(DateTime));
            Map(m => m.ReportDate).Name("Report Date").Index(1).Default(default(DateTime));
            Map(m => m.ReportPeriod).Name("Report Period").Index(2).Default(0);
            Map(m => m.IM).Name("IM").Index(3).Default(0);
            Map(m => m.AudioVideo).Name("Audio/Video").Index(4).Default(0);
            Map(m => m.AppSharing).Name("App Sharing").Index(5).Default(0);
            Map(m => m.Web).Name("Web").Index(6).Default(0);
            Map(m => m.DialInOut3rdParty).Name("Dial-in/out 3rd Party").Index(7).Default(0);
        }
    }

    /// <summary>
    /// Get usage trends on the number of unique users and type of conference sessions that users from your organization participated in. Types of conference sessions include IM, audio/video, application sharing, web, and dial-in/out - 3rd party
    /// </summary>
    public class SkypeForBusinessParticipantActivityUserCounts : JSONODataBase
    {
        [JsonProperty("reportRefreshDate")]
        public DateTime ReportRefreshDate { get; set; }

        [JsonProperty("reportDate")]
        public DateTime ReportDate { get; set; }

        [JsonProperty("reportPeriod")]
        public int ReportPeriod { get; set; }

        [JsonProperty("im")]
        public Nullable<Int64> IM { get; set; }

        [JsonProperty("audioVideo")]
        public Nullable<Int64> AudioVideo { get; set; }

        [JsonProperty("appSharing")]
        public Nullable<Int64> AppSharing { get; set; }

        [JsonProperty("web")]
        public Nullable<Int64> Web { get; set; }

        [JsonProperty("dialInOut3rdParty")]
        public Nullable<Int64> DialInOut3rdParty { get; set; }


    }
}
