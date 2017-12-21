using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Reports.o365Graph.TenantReport
{
    public class OneDriveActivityDetail : JSONODataBase
    {
        [JsonProperty("reportRefreshDate")]
        public DateTime ReportRefreshDate { get; set; }

        [JsonProperty("lastActivityDate")]
        public DateTime LastActivityDateUTC { get; set; }

        [JsonProperty("userPrincipalName")]
        public string UPN { get; set; }

        [JsonProperty("isDeleted")]
        public string Deleted { get; set; }

        [JsonProperty("deletedDate")]
        public Nullable<DateTime> DeletedDate { get; set; }

        [JsonProperty("viewedOrEditedFileCount")]
        public Int32 FilesViewedModified { get; set; }

        [JsonProperty("syncedFileCount")]
        public Int32 SyncedFileCount { get; set; }

        [JsonProperty("sharedInternallyFileCount")]
        public Int64 SharedInternallyFileCount { get; set; }

        [JsonProperty("sharedExternallyFileCount")]
        public Int64 SharedExternallyFileCount { get; set; }

        [JsonProperty("assignedProducts")]
        public IEnumerable<string> ProductsAssigned { get; set; }

        [JsonIgnore()]
        public string ProductsAssignedCsv { get; set; }

        [JsonProperty("reportPeriod")]
        public int ReportPeriod { get; set; }


    }
}
