using CsvHelper.Configuration;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Reports.o365Graph.TenantReport.Mappings
{
    /*
     * CSV Mapping for User activity
     * Report Refresh Date,,,,,,,,Report Period
     * */
    internal class SkypeForBusinessDeviceUsageUserDetailMap : ClassMap<SkypeForBusinessDeviceUsageUserDetail>
    {
        internal SkypeForBusinessDeviceUsageUserDetailMap()
        {
            Map(m => m.ReportRefreshDate).Name("Report Refresh Date").Index(0).Default(default(DateTime));
            Map(m => m.UPN).Name("User Principal Name").Index(1).Default(string.Empty);
            Map(m => m.LastActivityDate).Name("Last Activity Date").Index(2).Default(default(Nullable<DateTime>));
            Map(m => m.UsedWindows).Name("Used Windows").Index(3).Default("No");
            Map(m => m.UsedWindowsPhone).Name("Used Windows Phone").Index(4).Default("No");
            Map(m => m.UsedAndroidPhone).Name("Used Android Phone").Index(5).Default("No");
            Map(m => m.UsediPhone).Name("Used iPhone").Index(6).Default("No");
            Map(m => m.UsediPad).Name("Used iPad").Index(7).Default("No");
            Map(m => m.ReportPeriod).Name("Report Period").Index(8).Default(0);
        }
    }


    /// <summary>
    /// Get details about Skype for Business device usage by user.
    /// </summary>
    public class SkypeForBusinessDeviceUsageUserDetail : JSONODataBase
    {
        [JsonProperty("reportRefreshDate")]
        public DateTime ReportRefreshDate { get; set; }

        [JsonProperty("reportPeriod")]
        public int ReportPeriod { get; set; }

        [JsonProperty("userPrincipalName")]
        public string UPN { get; set; }

        [JsonProperty("lastActivityDate")]
        public Nullable<DateTime> LastActivityDate { get; set; }

        [JsonProperty("usedWindows")]
        public string UsedWindows { get; set; }

        [JsonProperty("usedWindowsPhone")]
        public string UsedWindowsPhone { get; set; }

        [JsonProperty("usedAndroidPhone")]
        public string UsedAndroidPhone { get; set; }

        [JsonProperty("usediPhone")]
        public string UsediPhone { get; set; }

        [JsonProperty("usediPad")]
        public string UsediPad { get; set; }
    }
}
