using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Reports.o365Graph.TenantReport
{
    /// <summary>
    /// WebResponse:
    ///     Data as of,ExchangeActive,ExchangeInActive,OneDriveActive,OneDriveInActive,SharePointActive,SharePointInActive,SkypeActive,SkypeInActive,YammerActive,YammerInActive,Reporting period in days
    /// </summary>
    public class Office365ActiveUsersServices
    {
        public DateTime DataAsOf { get; set; }
        public long ExchangeActive { get; set; }
        public long ExchangeInActive { get; set; }
        public long OneDriveActive { get; set; }
        public long OneDriveInActive { get; set; }
        public long SharePointActive { get; set; }
        public long SharePointInActive { get; set; }
        public long SkypeActive { get; set; }
        public long SkypeInActive { get; set; }
        public long YammerActive { get; set; }
        public long YammerInActive { get; set; }
        public int ReportingPeriodDays { get; set; }
    }
}
