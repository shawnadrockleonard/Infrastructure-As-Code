using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Reports.o365Graph.TenantReport
{
    public class OneDriveActivityDetail
    {
        /*
         * 
         * 
Data as of,
User principal name,
Deleted,
Deleted date,
Last activity date (UTC),
Files viewed or edited,
Files synced,
Files shared internally,
Files shared externally,
Products assigned,
Reporting period in days
         */

        public DateTime DataAsOf { get; set; }

        public DateTime LastActivityDateUTC { get; set; }

        public string UPN { get; set; }

        public string Deleted { get; set; }

        public Nullable<DateTime> DeletedDate { get; set; }

        public Int32 FilesViewedModified { get; set; }

        public Int32 FilesSynced { get; set; }

        public Int64 FilesSharedINT { get; set; }

        public Int64 FilesSharedEXT { get; set; }

        public string ProductsAssigned { get; set; }

        public int ReportingPeriodDays { get; set; }
    }
}
