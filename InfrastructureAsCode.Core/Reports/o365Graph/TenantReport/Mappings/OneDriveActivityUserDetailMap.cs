using CsvHelper.Configuration;
using InfrastructureAsCode.Core.Reports.o365Graph.TenantReport;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Reports.o365Graph.TenantReport.Mappings
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
    internal class OneDriveActivityUserDetailMap : ClassMap<OneDriveActivityUserDetail>
    {
        internal OneDriveActivityUserDetailMap()
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
            Map(m => m.ProductsAssignedCSV).Name("Assigned Products").Index(9).Default(string.Empty);
            Map(m => m.ReportPeriod).Name("Report Period").Index(10).Default(0);
        }
    }


    public class OneDriveActivityUserDetail : JSONODataBase
    {
        [JsonProperty("reportRefreshDate")]
        public DateTime ReportRefreshDate { get; set; }

        [JsonProperty("lastActivityDate")]
        public DateTime LastActivityDateUTC { get; set; }

        [JsonProperty("userPrincipalName")]
        public string UPN { get; set; }

        [JsonProperty("isDeleted")]
        public string Deleted { get; set; }

        [JsonProperty("deletedDate")]
        public Nullable<DateTime> DeletedDate { get; set; }

        [JsonProperty("viewedOrEditedFileCount")]
        public Nullable<Int64> FilesViewedModified { get; set; }

        [JsonProperty("syncedFileCount")]
        public Nullable<Int64> SyncedFileCount { get; set; }

        [JsonProperty("sharedInternallyFileCount")]
        public Nullable<Int64> SharedInternallyFileCount { get; set; }

        [JsonProperty("sharedExternallyFileCount")]
        public Nullable<Int64> SharedExternallyFileCount { get; set; }

        [JsonProperty("assignedProducts")]
        public IEnumerable<string> ProductsAssigned { get; set; }

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
