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
     * 
Report Refresh Date,
Visited Page Count,
Report Date,
Report Period
 */
    class SharePointActivityPagesMap : ClassMap<SharePointActivityPages>
    {
        public SharePointActivityPagesMap()
        {
            Map(m => m.ReportRefreshDate).Name("Report Refresh Date").Index(0).Default(default(DateTime));
            Map(m => m.VisitedPageCount).Name("Visited Page Count").Index(1).Default(0);
            Map(m => m.ReportDate).Name("Report Date").Index(2).Default(default(DateTime));
            Map(m => m.ReportPeriod).Name("Report Period").Index(3).Default(0);
        }
    }


    public class SharePointActivityPages : JSONODataBase
    {
        [JsonProperty("reportRefreshDate")]
        public DateTime ReportRefreshDate { get; set; }

        [JsonProperty("visitedPageCount")]
        public Int64 VisitedPageCount { get; set; }

        [JsonProperty("reportDate")]
        public DateTime ReportDate { get; set; }

        [JsonProperty("reportPeriod")]
        public int ReportPeriod { get; set; }
    }
}
