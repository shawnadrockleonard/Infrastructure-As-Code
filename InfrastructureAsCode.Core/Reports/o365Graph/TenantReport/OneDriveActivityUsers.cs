using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Reports.o365Graph.TenantReport
{
    public class OneDriveActivityUsers
    {
        /*
Data as of,
Viewed or edited,
Synced,
Shared internally,
Shared externally,
Last activity date (UTC),
Reporting period in days
         */

        public DateTime DataAsOf { get; set; }

        public Int32 FilesViewedModified { get; set; }

        public Int32 FilesSynced { get; set; }

        public Int64 FilesSharedINT { get; set; }

        public Int64 FilesSharedEXT { get; set; }

        public DateTime LastActivityDateUTC { get; set; }

        public int ReportingPeriodDays { get; set; }
    }
}
