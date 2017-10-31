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
User Principal Name,
Is Deleted,
Deleted Date,
Last Activity Date,
Viewed Or Edited File Count,
Synced File Count,
Shared Internally File Count,
Shared Externally File Count,
Visited Page Count,
Assigned Products,
Report Period
2017-10-28,<user>,False,,2017-10-28,25,0,1,0,4,E3,7
 */
    public class SharePointActivityUserDetailMap : ClassMap<SharePointActivityUserDetail>
    {
        public SharePointActivityUserDetailMap()
        {
            Map(m => m.ReportRefreshDate).Name("Report Refresh Date").Index(0).Default(default(DateTime));
            Map(m => m.UserPrincipalName).Name("User Principal Name").Index(1).Default(string.Empty);
            Map(m => m.IsDeleted).Name("Is Deleted").Index(2).Default(false);
            Map(m => m.DeletedDate).Name("Deleted Date").Index(3).Default(default(Nullable<DateTime>));
            Map(m => m.LastActivityDate).Name("Last Activity Date").Index(4).Default(default(DateTime));
            Map(m => m.ViewedOrEditedFileCount).Name("Viewed Or Edited File Count").Index(5).Default(0);
            Map(m => m.SyncedFileCount).Name("Synced File Count").Index(6).Default(0);
            Map(m => m.SharedInternallyFileCount).Name("Shared Internally File Count").Index(7).Default(0);
            Map(m => m.SharedExternallyFileCount).Name("Shared Externally File Count").Index(8).Default(0);
            Map(m => m.VisitedPageCount).Name("Visited Page Count").Index(9).Default(0);
            Map(m => m.AssignedProducts).Name("Assigned Products").Index(10).Default(string.Empty);
            Map(m => m.ReportPeriod).Name("Report Period").Index(11).Default(0);
        }
    }
}
