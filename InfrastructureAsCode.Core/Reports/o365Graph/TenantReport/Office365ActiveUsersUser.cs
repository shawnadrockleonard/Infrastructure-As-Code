using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Reports.o365Graph.TenantReport
{
    /// <summary>
    /// WebResponse: 
    ///     Data as of,Office 365,Exchange,OneDrive,SharePoint,Skype For Business,Yammer,Last activity date(UTC),Reporting period in days
    /// </summary>
    public class Office365ActiveUsersUser
    {
        public DateTime DataAsOf { get; set; }
        public long Office365 { get; set; }
        public long Exchange { get; set; }
        public long OneDrive { get; set; }
        public long SharePoint { get; set; }
        public long SkypeForBusiness { get; set; }
        public long Yammer { get; set; }
        public DateTime LastActivityDateUTC { get; set; }
        public int ReportingPeriodDays { get; set; }
    }
}
