using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Reports.o365Graph.TenantReport
{
    public class OneDriveActivityUserCounts
    {

        public DateTime ReportRefreshDate { get; set; }

        public Int32 FilesViewedModified { get; set; }

        public Int32 FilesSynced { get; set; }

        public Int64 FilesSharedINT { get; set; }

        public Int64 FilesSharedEXT { get; set; }

        public DateTime ReportDate { get; set; }

        public int ReportPeriod { get; set; }
    }
}
