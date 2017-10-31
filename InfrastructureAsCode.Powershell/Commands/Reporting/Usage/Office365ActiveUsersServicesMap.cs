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
Report Refresh Date
Exchange Active
Exchange Inactive
OneDrive Active
OneDrive Inactive
SharePoint Active
SharePoint Inactive
Skype For Business Active
Skype For Business Inactive
Yammer Active
Yammer Inactive
Teams Active
Teams Inactive
Report Period
     */
    class Office365ActiveUsersServicesMap : ClassMap<Office365ActiveUsersServices>
    {
        public Office365ActiveUsersServicesMap()
        {
            Map(m => m.ReportRefreshDate).Name("Report Refresh Date").Index(0).Default(default(DateTime));
            Map(m => m.ExchangeActive).Name("Exchange Active").Index(1).Default(0);
            Map(m => m.ExchangeInActive).Name("Exchange Inactive").Index(2).Default(0);
            Map(m => m.OneDriveActive).Name("OneDrive Active").Index(3).Default(0);
            Map(m => m.OneDriveInActive).Name("OneDrive Inactive").Index(4).Default(0);
            Map(m => m.SharePointActive).Name("SharePoint Active").Index(5).Default(0);
            Map(m => m.SharePointInActive).Name("SharePoint Inactive").Index(6).Default(0);
            Map(m => m.SkypeActive).Name("Skype For Business Active").Index(7).Default(0);
            Map(m => m.SkypeInActive).Name("Skype For Business Inactive").Index(8).Default(0);
            Map(m => m.YammerActive).Name("Yammer Active").Index(9).Default(0);
            Map(m => m.YammerInActive).Name("Yammer Inactive").Index(10).Default(0);
            Map(m => m.MSTeamActive).Name("Teams Active").Index(11).Default(0);
            Map(m => m.MSTeamInActive).Name("Teams Inactive").Index(12).Default(0);
            Map(m => m.ReportingPeriodDays).Name("Report Period").Index(13).Default(0);
        }
    }


}
