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
The CSV file has the following headers for columns.
Report Refresh Date
Site Type
Storage Used (Byte)
Report Date
Report Period
     */
    public class OneDriveUsageStorageMap : ClassMap<OneDriveUsageStorage>
    {
        public OneDriveUsageStorageMap()
        {
            Map(m => m.ReportRefreshDate).Name("Report Refresh Date").Index(0).Default(default(DateTime));
            Map(m => m.SiteType).Name("Site Type").Index(1).Default(string.Empty);
            Map(m => m.StorageUsed_Bytes).Name("Storage Used (Byte)").Index(2).Default(0);
            Map(m => m.ReportDate).Name("Report Date").Index(3).Default(default(DateTime));
            Map(m => m.ReportingPeriodDays).Name("Report Period").Index(4).Default(0);
        }
    }
}
