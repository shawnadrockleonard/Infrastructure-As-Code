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
Report Refresh Date
Office 365
Exchange
OneDrive
SharePoint
Skype For Business
Yammer
Teams
Report Date
Report Period
     */
    class Office365ActiveUsersUserCountsMap : ClassMap<Office365ActiveUsersUserCounts>
    {
        public Office365ActiveUsersUserCountsMap()
        {
            Map(m => m.ReportRefreshDate).Name("Report Refresh Date").Index(0).Default(default(DateTime));
            Map(m => m.Office365).Name("Office 365").Index(1).Default(0);
            Map(m => m.Exchange).Name("Exchange").Index(2).Default(0);
            Map(m => m.OneDrive).Name("OneDrive").Index(3).Default(0);
            Map(m => m.SharePoint).Name("SharePoint").Index(4).Default(0);
            Map(m => m.SkypeForBusiness).Name("Skype For Business").Index(5).Default(0);
            Map(m => m.Yammer).Name("Yammer").Index(6).Default(0);
            Map(m => m.Yammer).Name("Teams").Index(7).Default(0);
            Map(m => m.ReportDate).Name("Report Date").Index(8).Default(default(DateTime));
            Map(m => m.ReportingPeriodDays).Name("Report Period").Index(9).Default(0);
        }
    }


    public class Office365ActiveUsersUserCounts : JSONODataBase
    {
        [JsonProperty("reportRefreshDate")]
        public DateTime ReportRefreshDate { get; set; }

        [JsonProperty("office365")]
        public Nullable<Int64> Office365 { get; set; }

        [JsonProperty("exchange")]
        public Nullable<Int64> Exchange { get; set; }

        [JsonProperty("oneDrive")]
        public Nullable<Int64> OneDrive { get; set; }

        [JsonProperty("sharePoint")]
        public Nullable<Int64> SharePoint { get; set; }

        [JsonProperty("skypeForBusiness")]
        public Nullable<Int64> SkypeForBusiness { get; set; }

        [JsonProperty("yammer")]
        public Nullable<Int64> Yammer { get; set; }

        [JsonProperty("teams")]
        public Nullable<Int64> MSTeams { get; set; }

        [JsonProperty("reportDate")]
        public DateTime ReportDate { get; set; }

        [JsonProperty("reportPeriod")]
        public int ReportingPeriodDays { get; set; }
    }
}