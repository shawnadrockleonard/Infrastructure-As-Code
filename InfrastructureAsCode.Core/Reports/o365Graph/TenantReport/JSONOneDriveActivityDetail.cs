using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Reports.o365Graph.TenantReport
{
    public class JSONOneDriveActivityDetail : JSONODataBase
    {

        [JsonProperty("reportRefreshDate")]
        public DateTime ReportRefreshDate { get; set; }

        [JsonProperty("viewedOrEdited")]
        public Nullable<long> FilesViewedModified { get; set; }

        [JsonProperty("synced")]
        public Nullable<long> FilesSynced { get; set; }

        [JsonProperty("sharedInternally")]
        public Nullable<long> FilesSharedINT { get; set; }

        [JsonProperty("sharedExternally")]
        public Nullable<long> FilesSharedEXT { get; set; }

        [JsonProperty("reportDate")]
        public DateTime ReportDate { get; set; }

        [JsonProperty("reportPeriod")]
        public Int32 ReportPeriod { get; set; }

    }
}
