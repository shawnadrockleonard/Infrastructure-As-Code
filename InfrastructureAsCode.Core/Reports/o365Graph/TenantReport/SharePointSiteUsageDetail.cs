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

        public Int64 FileCount { get; set; }

        public Int64 ActiveFileCount { get; set; }

        public Int64 PageViewCount { get; set; }

        public Int64 VisitedPageCount { get; set; }

        public Int64 StorageUsed_Byte { get; set; }

        public Int64 StorageAllocated_Byte { get; set; }

        public string RootWebTemplate { get; set; }

        public int ReportPeriod { get; set; }
    }
}
