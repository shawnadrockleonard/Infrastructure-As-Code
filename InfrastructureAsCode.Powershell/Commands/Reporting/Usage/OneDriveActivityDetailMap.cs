using CsvHelper.Configuration;
using InfrastructureAsCode.Core.Reports.o365Graph.TenantReport;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Powershell.Commands.Reporting.Usage
{
    /*
     * 
The CSV file has the following headers for columns.
Report Refresh Date,
User Principal Name,
Is Deleted,
Deleted Date,
Last Activity Date,
Viewed Or Edited File Count,
Synced File Count,
Shared Internally File Count,
Shared Externally File Count,
Assigned Products,
Report Period
     */
    public class OneDriveActivityDetailMap : ClassMap<OneDriveActivityDetail>
    {
        public OneDriveActivityDetailMap()
        {
            Map(m => m.ReportRefreshDate).Name("Report Refresh Date").Index(0).Default(default(DateTime));
            Map(m => m.UPN).Name("User Principal Name").Index(1).Default(string.Empty);
            Map(m => m.Deleted).Name("Is Deleted").Index(2).Default(string.Empty);
            Map(m => m.DeletedDate).Name("Deleted Date").Index(3).Default(default(Nullable<DateTime>));
            Map(m => m.LastActivityDateUTC).Name("Last Activity Date").Index(4).Default(default(DateTime));
            Map(m => m.FilesViewedModified).Name("Viewed Or Edited File Count").Index(5).Default(0);
            Map(m => m.SyncedFileCount).Name("Synced File Count").Index(6).Default(0);
            Map(m => m.SharedInternallyFileCount).Name("Shared Internally File Count").Index(7).Default(0);
            Map(m => m.SharedExternallyFileCount).Name("Shared Externally File Count").Index(8).Default(0);
            Map(m => m.ProductsAssigned).Name("Assigned Products").Index(9).Default(string.Empty);
            Map(m => m.ReportPeriod).Name("Report Period").Index(10).Default(0);
        }
    }
}
