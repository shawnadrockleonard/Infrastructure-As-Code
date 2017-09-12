using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Reports.o365Graph.TenantReport
{
    public class OneDriveUsageDetail
    {
        /*
         * 
         * 
Data as of,
Site URL,
Site owner,
Deleted,
Last activity date (UTC),
Files,
Files viewed or edited,
Storage used (B),
Storage allocated (B),
Reporting period in days
         */

        public DateTime DataAsOf { get; set; }

        public DateTime LastActivityDateUTC { get; set; }

        public string SiteURL { get; set; }

        public string SiteOwner { get; set; }

        public string Deleted { get; set; }

        public Int32 Files { get; set; }

        public Int32 FilesViewedModified { get; set; }

        public Int64 Storage_Used_B { get; set; }

        public Int64 Storage_Allocated_B { get; set; }

        public int ReportingPeriodDays { get; set; }
    }
}
