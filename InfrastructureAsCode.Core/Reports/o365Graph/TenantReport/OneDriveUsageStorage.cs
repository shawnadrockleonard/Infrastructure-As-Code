using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Reports.o365Graph.TenantReport
{
    public class OneDriveUsageStorage
    {
        /*
Data as of,
SiteType,
Storage used,
Last activity date (UTC),
Reporting period in days
         */

        public DateTime DataAsOf { get; set; }

        public DateTime LastActivityDateUTC { get; set; }

        public string SiteType { get; set; }

        public Int64 Storage_Used_B { get; set; }

        public int ReportingPeriodDays { get; set; }
    }
}
