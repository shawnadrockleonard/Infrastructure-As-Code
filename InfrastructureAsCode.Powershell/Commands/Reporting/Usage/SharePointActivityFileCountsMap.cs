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
Report Refresh Date,
Viewed Or Edited,
Synced,
Shared Internally,
Shared Externally,
Report Date,
Report Period
     */
    class SharePointActivityFileCountsMap : ClassMap<SharePointActivityFileCounts>
    {
        public SharePointActivityFileCountsMap()
        {
            Map(m => m.ReportRefreshDate).Name("Report Refresh Date").Index(0).Default(default(DateTime));
            Map(m => m.ViewedOrEdited).Name("Viewed Or Edited").Index(1).Default(0);
            Map(m => m.Synced).Name("Synced").Index(2).Default(0);
            Map(m => m.SharedInternally).Name("Shared Internally").Index(3).Default(0);
            Map(m => m.SharedExternally).Name("Shared Externally").Index(4).Default(0);
            Map(m => m.ReportDate).Name("Report Date").Index(5).Default(default(DateTime));
            Map(m => m.ReportPeriod).Name("Report Period").Index(6).Default(0);
        }
    }
}
