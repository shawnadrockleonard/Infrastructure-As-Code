using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Reports.o365Graph.TenantReport
{
    public class OneDriveUsageAccount
    {
        /*
Data as of,
SiteType,
Total accounts,
Active accounts,
Last activity date (UTC),
Reporting period in days
         */

        public DateTime DataAsOf { get; set; }

        public DateTime LastActivityDateUTC { get; set; }

        public string SiteType { get; set; }

        public Int64 Total_Accounts { get; set; }

        public Int64 Active_Accounts { get; set; }

        public int ReportingPeriodDays { get; set; }
    }
}
