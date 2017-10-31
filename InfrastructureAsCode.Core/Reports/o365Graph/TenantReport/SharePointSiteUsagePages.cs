using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Reports.o365Graph.TenantReport
{
    public class SharePointSiteUsagePages
    {
        public DateTime ReportRefreshDate { get; set; }

        public string SiteType { get; set; }

        public Int64 PageViewCount { get; set; }

        public DateTime ReportDate { get; set; }

        public int ReportPeriod { get; set; }
    }
}
