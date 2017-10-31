using InfrastructureAsCode.Core.Reports.o365Graph.TenantReport;
using CsvHelper.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Powershell.Commands.Reporting.Usage
{
    /*
     * 
     * 
Report Refresh Date,
Visited Page Count,
Report Date,
Report Period
 */
    class SharePointActivityPagesMap : ClassMap<SharePointActivityPages>
    {
        public SharePointActivityPagesMap()
        {
            Map(m => m.ReportRefreshDate).Name("Report Refresh Date").Index(0).Default(default(DateTime));
            Map(m => m.VisitedPageCount).Name("Visited Page Count").Index(1).Default(0);
            Map(m => m.ReportDate).Name("Report Date").Index(2).Default(default(DateTime));
            Map(m => m.ReportPeriod).Name("Report Period").Index(3).Default(0);
        }
    }
}
