using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Reports.o365Graph.TenantReport
{
    /// <summary>
    /// WebResponse:
    ///     DataAsOf,UPN,DisplayName,Deleted,DeletedDate,
    ///         LicenseForExchange,LicenseForOneDrive,LicenseForSharePoint,LicenseForSkypeForBusiness,LicenseForYammer,
    ///         LastActivityDateForExchange,LastActivityDateForOneDrive,LastActivityDateForSharePoint,LastActivityDateForSkypeForBusiness,LastActivityDateForYammer
    ///         LicenseAssignedDateForExchange,LicenseAssignedDateForOneDrive,LicenseAssignedDateForSharePoint,LicenseAssignedDateForSkypeForBusiness,LicenseAssignedDateForYammer
    ///         ProductsAssigned
    /// </summary>
    public class Office365ActiveUsersDetails
    {
        public DateTime DataAsOf { get; set; }
        public string UPN { get; set; }
        public string DisplayName { get; set; }
        public bool Deleted { get; set; }
        public Nullable<DateTime> DeletedDate { get; set; }
        public bool LicenseForExchange { get; set; }
        public bool LicenseForOneDrive { get; set; }
        public bool LicenseForSharePoint { get; set; }
        public bool LicenseForSkypeForBusiness { get; set; }
        public bool LicenseForYammer { get; set; }
        public Nullable<DateTime> LastActivityDateForExchange { get; set; }
        public Nullable<DateTime> LastActivityDateForOneDrive { get; set; }
        public Nullable<DateTime> LastActivityDateForSharePoint { get; set; }
        public Nullable<DateTime> LastActivityDateForSkypeForBusiness { get; set; }
        public Nullable<DateTime> LastActivityDateForYammer { get; set; }
        public Nullable<DateTime> LicenseAssignedDateForExchange { get; set; }
        public Nullable<DateTime> LicenseAssignedDateForOneDrive { get; set; }
        public Nullable<DateTime> LicenseAssignedDateForSharePoint { get; set; }
        public Nullable<DateTime> LicenseAssignedDateForSkypeForBusiness { get; set; }
        public Nullable<DateTime> LicenseAssignedDateForYammer { get; set; }
        public string ProductsAssigned { get; set; }
    }
}
