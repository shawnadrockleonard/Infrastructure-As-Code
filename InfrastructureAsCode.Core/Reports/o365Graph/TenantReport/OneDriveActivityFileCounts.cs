using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Reports.o365Graph.TenantReport
{
    public class OneDriveActivityFileCounts : JSONODataBase
    {
        [JsonProperty("reportRefreshDate")]
        public DateTime ReportRefreshDate { get; set; }

        [JsonProperty("viewedOrEdited")]
        public Nullable<Int32> FilesViewedModified { get; set; }

        [JsonProperty("synced")]
        public Nullable<Int32> FilesSynced { get; set; }

        [JsonProperty("sharedInternally")]
        public Nullable<Int64> FilesSharedINT { get; set; }

        [JsonProperty("sharedExternally")]
        public Nullable<Int64> FilesSharedEXT { get; set; }

        [JsonProperty("reportDate")]
        public DateTime ReportDate { get; set; }

        [JsonProperty("reportPeriod")]
        public int ReportPeriod { get; set; }
    }
}
