using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Reports.o365Graph.TenantReport
{
    public class OneDriveUsageAccountCounts : JSONODataBase
    {
        [JsonProperty("reportRefreshDate")]
        public DateTime ReportRefreshDate { get; set; }

        [JsonProperty("reportDate")]
        public DateTime ReportDate { get; set; }

        [JsonProperty("siteType")]
        public string SiteType { get; set; }

        [JsonProperty("total")]
        public Int64 Total_Accounts { get; set; }

        [JsonProperty("active")]
        public Int64 Active_Accounts { get; set; }

        [JsonProperty("reportPeriod")]
        public int ReportPeriod { get; set; }

    }
}
