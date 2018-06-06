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
    class SkypeForBusinessOrganizerActivityMinuteCountsMap : ClassMap<SkypeForBusinessOrganizerActivityMinuteCounts>
    {
        SkypeForBusinessOrganizerActivityMinuteCountsMap()
        {
            Map(m => m.ReportRefreshDate).Name("Report Refresh Date").Index(0).Default(default(DateTime));
            Map(m => m.ReportDate).Name("Report Date").Index(1).Default(default(Nullable<DateTime>));
            Map(m => m.ReportPeriod).Name("Report Period").Index(2).Default(0);
            Map(m => m.AudioVideo).Name("Audio/Video").Index(3).Default(0);
            Map(m => m.DialInMicrosoft).Name("Dial-in Microsoft").Index(4).Default(0);
            Map(m => m.DialOutMicrosoft).Name("Dial-out Microsoft").Index(5).Default(0);
        }
    }

    /// <summary>
    /// Get usage trends on the length in minutes and type of conference sessions held and organized by users in your organization. Types of conference sessions include audio/video, and dial-in and dial-out - Microsoft.
    /// </summary>
    public class SkypeForBusinessOrganizerActivityMinuteCounts : JSONODataBase
    {
        [JsonProperty("reportRefreshDate")]
        public DateTime ReportRefreshDate { get; set; }

        [JsonProperty("reportDate")]
        public Nullable<DateTime> ReportDate { get; set; }

        [JsonProperty("reportPeriod")]
        public int ReportPeriod { get; set; }

        [JsonProperty("audioVideo")]
        public Nullable<Int64> AudioVideo { get; set; }

        [JsonProperty("dialInMicrosoft")]
        public Nullable<Int64> DialInMicrosoft { get; set; }

        [JsonProperty("dialOutMicrosoft")]
        public Nullable<Int64> DialOutMicrosoft { get; set; }

    }
}
