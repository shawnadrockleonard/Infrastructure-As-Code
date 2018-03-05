using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Reports.o365Graph.TenantReport
{
    public class SharePointSiteUsageDetail
    {
        public DateTime ReportRefreshDate { get; set; }

        public string SiteURL { get; set; }

        public string OwnerDisplayName { get; set; }

        public bool IsDeleted { get; set; }

        public Nullable<DateTime> LastActivityDate { get; set; }

        public long FileCount { get; set; }

        public long ActiveFileCount { get; set; }

        public long PageViewCount { get; set; }

        public long VisitedPageCount { get; set; }

        public long StorageUsed_Byte { get; set; }

        public long StorageAllocated_Byte { get; set; }

        public string RootWebTemplate { get; set; }

        public int ReportPeriod { get; set; }
    }
}
