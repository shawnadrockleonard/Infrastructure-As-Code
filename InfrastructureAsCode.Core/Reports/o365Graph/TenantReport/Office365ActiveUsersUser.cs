using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Reports.o365Graph.TenantReport
{
    public class Office365ActiveUsersUser
    {
        public DateTime ReportRefreshDate { get; set; }
        public long Office365 { get; set; }
        public long Exchange { get; set; }
        public long OneDrive { get; set; }
        public long SharePoint { get; set; }
        public long SkypeForBusiness { get; set; }
        public long Yammer { get; set; }
        public long MSTeams { get; set; }
        public DateTime ReportDate { get; set; }
        public int ReportingPeriodDays { get; set; }
    }
}
