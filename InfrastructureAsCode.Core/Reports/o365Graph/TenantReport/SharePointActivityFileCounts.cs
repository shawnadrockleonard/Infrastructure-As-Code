using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Reports.o365Graph.TenantReport
{
    /*
     * 

     */
    public class SharePointActivityFileCounts
    {
        public DateTime ReportRefreshDate { get; set; }


        public Int64 ViewedOrEdited { get; set; }

        public Int64 Synced { get; set; }

        public Int64 SharedInternally { get; set; }

        public Int64 SharedExternally { get; set; }

        public DateTime ReportDate { get; set; }

        public int ReportPeriod { get; set; }
    }
}
