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
     * CSV mapping for v1.0 and beta endpoints
     * Report Refresh Date,Report Date,Report Period,Peer-to-peer,Organized,Participated
     */
    internal class SkypeForBusinessActivityActivityCountsMap : ClassMap<SkypeForBusinessActivityActivityCounts>
    {
        internal SkypeForBusinessActivityActivityCountsMap()
        {
            Map(m => m.ReportRefreshDate).Name("Report Refresh Date").Index(0).Default(default(DateTime));
            Map(m => m.ReportDate).Name("Report Date").Index(1).Default(default(DateTime));
            Map(m => m.ReportPeriod).Name("Report Period").Index(2).Default(0);
            Map(m => m.PeerToPeer).Name("Peer-to-peer").Index(3).Default(0);
            Map(m => m.Organized).Name("Organized").Index(4).Default(0);
            Map(m => m.Participated).Name("Participated").Index(5).Default(0);
        }
    }

    /// <summary>
    /// Get the trends on how many users organized and participated in conference sessions held in your organization through Skype for Business. The report also includes the number of peer-to-peer sessions.
    /// </summary>
    public class SkypeForBusinessActivityActivityCounts : JSONODataBase
    {
        [JsonProperty("reportRefreshDate")]
        public DateTime ReportRefreshDate { get; set; }

        [JsonProperty("reportDate")]
        public DateTime ReportDate { get; set; }

        [JsonProperty("reportPeriod")]
        public int ReportPeriod { get; set; }

        [JsonProperty("peerToPeer")]
        public Nullable<Int64> PeerToPeer { get; set; }

        [JsonProperty("organized")]
        public Nullable<Int64> Organized { get; set; }

        [JsonProperty("participated")]
        public Nullable<Int64> Participated { get; set; }
    }
}
