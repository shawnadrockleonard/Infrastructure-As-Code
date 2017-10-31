using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Reports.o365Graph.TenantReport
{

    public class SharePointActivityUserDetail
    {
        public DateTime ReportRefreshDate { get; set; }

        public string UserPrincipalName { get; set; }
        public bool IsDeleted { get; set; }
        public Nullable<DateTime> DeletedDate { get; set; }
        public DateTime LastActivityDate { get; set; }
        public Int64 ViewedOrEditedFileCount { get; set; }
        public Int64 SyncedFileCount { get; set; }
        public Int64 SharedInternallyFileCount { get; set; }
        public Int64 SharedExternallyFileCount { get; set; }
        public Int64 VisitedPageCount { get; set; }
        public string AssignedProducts { get; set; }

        public int ReportPeriod { get; set; }
    }
}
