using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Reports.o365Graph.TenantReport
{
    public class OneDriveUsageAccountDetail
    {
        public DateTime ReportRefreshDate { get; set; }

        public DateTime LastActivityDateUTC { get; set; }

        public string SiteURL { get; set; }

        public string SiteOwner { get; set; }

        public string Deleted { get; set; }

        public Int64 Files { get; set; }

        public Int64 FilesViewedModified { get; set; }

        public Int64 StorageUsed_Bytes { get; set; }

        public Int64 StorageAllocated_Bytes { get; set; }

        public int ReportingPeriodDays { get; set; }
    }
}
