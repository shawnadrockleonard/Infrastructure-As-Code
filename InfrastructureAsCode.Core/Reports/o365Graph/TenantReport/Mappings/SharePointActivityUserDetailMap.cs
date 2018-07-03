using InfrastructureAsCode.Core.Reports.o365Graph.TenantReport;
using CsvHelper.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace InfrastructureAsCode.Core.Reports.o365Graph.TenantReport.Mappings
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
    internal class SharePointActivityUserDetailMap : ClassMap<SharePointActivityUserDetail>
    {
        internal SharePointActivityUserDetailMap()
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
            Map(m => m.ProductsAssignedCSV).Name("Assigned Products").Index(10).Default(string.Empty);
            Map(m => m.ReportPeriod).Name("Report Period").Index(11).Default(0);
        }
    }



    /// <summary>
    /// Get details about SharePoint activity by user.
    /// </summary>
    public class SharePointActivityUserDetail : JSONODataBase
    {
        [JsonProperty("reportRefreshDate")]
        public DateTime ReportRefreshDate { get; set; }

        [JsonProperty("userPrincipalName")]
        public string UserPrincipalName { get; set; }

        [JsonProperty("isDeleted")]
        public bool IsDeleted { get; set; }

        [JsonProperty("deletedDate")]
        public Nullable<DateTime> DeletedDate { get; set; }

        [JsonProperty("lastActivityDate")]
        public DateTime LastActivityDate { get; set; }

        [JsonProperty("viewedOrEditedFileCount")]
        public Nullable<Int64> ViewedOrEditedFileCount { get; set; }

        [JsonProperty("syncedFileCount")]
        public Nullable<Int64> SyncedFileCount { get; set; }

        [JsonProperty("sharedInternallyFileCount")]
        public Nullable<Int64> SharedInternallyFileCount { get; set; }

        [JsonProperty("sharedExternallyFileCount")]
        public Nullable<Int64> SharedExternallyFileCount { get; set; }

        [JsonProperty("visitedPageCount")]
        public Nullable<Int64> VisitedPageCount { get; set; }

        [JsonProperty("assignedProducts")]
        public IEnumerable<string> ProductsAssigned { get; set; }

        /// <summary>
        /// Assigned Products separated by semicolon
        /// </summary>
        [JsonIgnore()]
        public string ProductsAssignedCSV { get; set; }


        /// <summary>
        /// Process the CSV or Array into a Delimited string
        /// </summary>
        [JsonIgnore()]
        public string RealizedProductsAssigned
        {
            get
            {
                var _productsAssigned = string.Empty;
                if (ProductsAssigned != null)
                {
                    _productsAssigned = string.Join(",", ProductsAssigned);
                }
                else if (!string.IsNullOrEmpty(ProductsAssignedCSV))
                {
                    _productsAssigned = ProductsAssignedCSV.Replace("+", ",");
                }

                return _productsAssigned;
            }
        }


        [JsonProperty("reportPeriod")]
        public int ReportPeriod { get; set; }
    }
}
