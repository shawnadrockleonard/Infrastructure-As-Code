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
    The CSV file has the following headers for columns.
    Report Refresh Date
    User Principal Name
Display Name
Is Deleted
Deleted Date
Has Exchange License
Has OneDrive License
Has SharePoint License
Has Skype For Business License
Has Yammer License
Has Teams License

Exchange Last Activity Date
OneDrive Last Activity Date
SharePoint Last Activity Date
Skype For Business Last Activity Date
Yammer Last Activity Date
Teams Last Activity Date

Exchange License Assign Date
OneDrive License Assign Date
SharePoint License Assign Date
Skype For Business License Assign Date
Yammer License Assign Date
Teams License Assign Date

Assigned Products
    */
    public class Office365ActiveUsersDetailsMap : ClassMap<Office365ActiveUsersDetails>
    {
        public Office365ActiveUsersDetailsMap()
        {
            Map(m => m.ReportRefreshDate).Name("Report Refresh Date").Index(0).Default(default(DateTime));
            Map(m => m.UPN).Name("User Principal Name").Index(1).Default(string.Empty);
            Map(m => m.DisplayName).Name("Display Name").Index(2).Default(string.Empty);
            Map(m => m.Deleted).Name("Is Deleted").Index(3).Default(false);
            Map(m => m.DeletedDate).Name("Deleted Date").Index(4).Default(default(Nullable<DateTime>));

            Map(m => m.LicenseForExchange).Name("Has Exchange License").Index(5).Default(false);
            Map(m => m.LicenseForOneDrive).Name("Has OneDrive License").Index(6).Default(false);
            Map(m => m.LicenseForSharePoint).Name("Has SharePoint License").Index(7).Default(false);
            Map(m => m.LicenseForSkypeForBusiness).Name("Has Skype For Business License").Index(8).Default(false);
            Map(m => m.LicenseForYammer).Name("Has Yammer License").Index(9).Default(false);
            Map(m => m.LicenseForMSTeams).Name("Has Teams License").Index(10).Default(false);

            Map(m => m.LastActivityDateForExchange).Name("Exchange Last Activity Date").Index(11).Default(default(Nullable<DateTime>));
            Map(m => m.LastActivityDateForOneDrive).Name("OneDrive Last Activity Date").Index(12).Default(default(Nullable<DateTime>));
            Map(m => m.LastActivityDateForSharePoint).Name("SharePoint Last Activity Date").Index(13).Default(default(Nullable<DateTime>));
            Map(m => m.LastActivityDateForSkypeForBusiness).Name("Skype For Business Last Activity Date").Index(14).Default(default(Nullable<DateTime>));
            Map(m => m.LastActivityDateForYammer).Name("Yammer Last Activity Date").Index(15).Default(default(Nullable<DateTime>));
            Map(m => m.LastActivityDateForMSTeams).Name("Teams Last Activity Date").Index(16).Default(default(Nullable<DateTime>));

            Map(m => m.LicenseAssignedDateForExchange).Name("Exchange License Assign Date").Index(17).Default(default(Nullable<DateTime>));
            Map(m => m.LicenseAssignedDateForOneDrive).Name("OneDrive License Assign Date").Index(18).Default(default(Nullable<DateTime>));
            Map(m => m.LicenseAssignedDateForSharePoint).Name("SharePoint License Assign Date").Index(19).Default(default(Nullable<DateTime>));
            Map(m => m.LicenseAssignedDateForSkypeForBusiness).Name("Skype For Business License Assign Date").Index(20).Default(default(Nullable<DateTime>));
            Map(m => m.LicenseAssignedDateForYammer).Name("Yammer License Assign Date").Index(21).Default(default(Nullable<DateTime>));
            Map(m => m.LicenseAssignedDateForMSTeams).Name("Teams License Assign Date").Index(22).Default(default(Nullable<DateTime>));

            Map(m => m.ProductsAssigned).Name("Assigned Products").Index(23).Default(string.Empty);
        }
    }

}
