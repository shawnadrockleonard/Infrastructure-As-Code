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
Viewed Or Edited,
Synced,
Shared Internally,
Shared Externally,
Report Date,
Report Period
     */
    internal class SharePointActivityFileCountsMap : ClassMap<SharePointActivityFileCounts>
    {
        internal SharePointActivityFileCountsMap()
        {
            Map(m => m.ReportRefreshDate).Name("Report Refresh Date").Index(0).Default(default(DateTime));
            Map(m => m.ViewedOrEdited).Name("Viewed Or Edited").Index(1).Default(0);
            Map(m => m.Synced).Name("Synced").Index(2).Default(0);
            Map(m => m.SharedInternally).Name("Shared Internally").Index(3).Default(0);
            Map(m => m.SharedExternally).Name("Shared Externally").Index(4).Default(0);
            Map(m => m.ReportDate).Name("Report Date").Index(5).Default(default(DateTime));
            Map(m => m.ReportPeriod).Name("Report Period").Index(6).Default(0);
        }
    }


    public class SharePointActivityFileCounts : JSONODataBase
    {
        [JsonProperty("reportRefreshDate")]
        public DateTime ReportRefreshDate { get; set; }

        [JsonProperty("viewedOrEdited")]
        public Nullable<Int64> ViewedOrEdited { get; set; }

        [JsonProperty("synced")]
        public Nullable<Int64> Synced { get; set; }

        [JsonProperty("sharedInternally")]
        public Nullable<Int64> SharedInternally { get; set; }

        [JsonProperty("sharedExternally")]
        public Nullable<Int64> SharedExternally { get; set; }

        [JsonProperty("reportDate")]
        public DateTime ReportDate { get; set; }

        [JsonProperty("reportPeriod")]
        public int ReportPeriod { get; set; }
    }
}
