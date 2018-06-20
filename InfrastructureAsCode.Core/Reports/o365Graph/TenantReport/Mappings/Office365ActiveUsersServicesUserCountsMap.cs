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
    class Office365ActiveUsersServicesUserCountsMap : ClassMap<Office365ActiveUsersServicesUserCounts>
    {
        public Office365ActiveUsersServicesUserCountsMap()
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


    public class Office365ActiveUsersServicesUserCounts : JSONODataBase
    {
        [JsonProperty("reportRefreshDate")]
        public DateTime ReportRefreshDate { get; set; }

        [JsonProperty("exchangeActive")]
        public Nullable<Int64> ExchangeActive { get; set; }

        [JsonProperty("exchangeInactive")]
        public Nullable<Int64> ExchangeInActive { get; set; }

        [JsonProperty("oneDriveActive")]
        public Nullable<Int64> OneDriveActive { get; set; }

        [JsonProperty("oneDriveInactive")]
        public Nullable<Int64> OneDriveInActive { get; set; }

        [JsonProperty("sharePointActive")]
        public Nullable<Int64> SharePointActive { get; set; }

        [JsonProperty("sharePointInactive")]
        public Nullable<Int64> SharePointInActive { get; set; }

        [JsonProperty("skypeForBusinessActive")]
        public Nullable<Int64> SkypeActive { get; set; }

        [JsonProperty("skypeForBusinessInactive")]
        public Nullable<Int64> SkypeInActive { get; set; }

        [JsonProperty("yammerActive")]
        public Nullable<Int64> YammerActive { get; set; }

        [JsonProperty("yammerInactive")]
        public Nullable<Int64> YammerInActive { get; set; }

        [JsonProperty("teamsActive")]
        public Nullable<Int64> MSTeamActive { get; set; }

        [JsonProperty("teamsInactive")]
        public Nullable<Int64> MSTeamInActive { get; set; }

        [JsonProperty("reportPeriod")]
        public int ReportingPeriodDays { get; set; }
    }
}
