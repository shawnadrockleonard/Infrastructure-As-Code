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
The CSV file has the following headers for columns.
Report Refresh Date
Viewed Or Edited
Synced
Shared Internally
Shared Externally
Report Date
Report Period
     */
    internal class OneDriveActivityUserCountsMap : ClassMap<OneDriveActivityUserCounts>
    {
        internal OneDriveActivityUserCountsMap()
        {
            Map(m => m.ReportRefreshDate).Name("Report Refresh Date").Index(0).Default(default(DateTime));
            Map(m => m.FilesViewedModified).Name("Viewed Or Edited").Index(1).Default(0);
            Map(m => m.FilesSynced).Name("Synced").Index(2).Default(0);
            Map(m => m.FilesSharedINT).Name("Shared Internally").Index(3).Default(0);
            Map(m => m.FilesSharedEXT).Name("Shared Externally").Index(4).Default(0);
            Map(m => m.ReportDate).Name("Report Date").Index(5).Default(default(DateTime));
            Map(m => m.ReportPeriod).Name("Report Period").Index(6).Default(0);
        }
    }


    public class OneDriveActivityUserCounts : JSONODataBase
    {

        [JsonProperty("reportRefreshDate")]
        public DateTime ReportRefreshDate { get; set; }

        [JsonProperty("viewedOrEdited")]
        public Nullable<Int64> FilesViewedModified { get; set; }

        [JsonProperty("synced")]
        public Nullable<Int64> FilesSynced { get; set; }

        [JsonProperty("sharedInternally")]
        public Nullable<Int64> FilesSharedINT { get; set; }

        [JsonProperty("sharedExternally")]
        public Nullable<Int64> FilesSharedEXT { get; set; }

        [JsonProperty("reportDate")]
        public DateTime ReportDate { get; set; }

        [JsonProperty("reportPeriod")]
        public int ReportPeriod { get; set; }

    }
}
