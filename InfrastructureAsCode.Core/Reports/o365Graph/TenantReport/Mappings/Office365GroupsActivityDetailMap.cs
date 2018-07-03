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
    * The CSV file has the following headers for columns.
    * Report Refresh Date,,,,,,,,,,,,,,,,,Report Period
     */
    internal class Office365GroupsActivityDetailMap : ClassMap<Office365GroupsActivityDetail>
    {
        internal Office365GroupsActivityDetailMap()
        {
            Map(m => m.ReportRefreshDate).Name("Report Refresh Date").Index(0).Default(default(DateTime));
            Map(m => m.GroupName).Name("Group Display Name").Index(1).Default(string.Empty);
            Map(m => m.IsDeleted).Name("Is Deleted").Index(2).Default(false);
            Map(m => m.OwnerPrincipalName).Name("Owner Principal Name").Index(3).Default(string.Empty);
            Map(m => m.LastActivityDate).Name("Last Activity Date").Index(4).Default(default(DateTime));
            Map(m => m.GroupType).Name("Group Type").Index(5).Default(string.Empty);
            Map(m => m.MemberCount).Name("Member Count").Index(6).Default(0);
            Map(m => m.ExternalMemberCount).Name("External Member Count").Index(7).Default(0);
            Map(m => m.ExchangeReceivedEmailCount).Name("Exchange Received Email Count").Index(8).Default(0);
            Map(m => m.SharePointActiveFileCount).Name("SharePoint Active File Count").Index(9).Default(0);
            Map(m => m.YammerPostedMessageCount).Name("Yammer Posted Message Count").Index(10).Default(0);
            Map(m => m.YammerReadMessageCount).Name("Yammer Read Message Count").Index(11).Default(0);
            Map(m => m.YammerLikedMessageCount).Name("Yammer Liked Message Count").Index(12).Default(0);
            Map(m => m.ExchangeMailboxTotalItemCount).Name("Exchange Mailbox Total Item Count").Index(13).Default(0);
            Map(m => m.ExchangeMailboxStorageUsed_Byte).Name("Exchange Mailbox Storage Used (Byte)").Index(14).Default(0);
            Map(m => m.SharePointTotalFileCount).Name("SharePoint Total File Count").Index(15).Default(0);
            Map(m => m.SharePointSiteStorageUsed_Byte).Name("SharePoint Site Storage Used (Byte)").Index(16).Default(0);
            Map(m => m.ReportPeriod).Name("Report Period").Index(17).Default(string.Empty);
        }
    }



    public class Office365GroupsActivityDetail : JSONODataBase
    {

        [JsonProperty("reportRefreshDate")]
        public DateTime ReportRefreshDate { get; set; }

        [JsonProperty("groupDisplayName")]
        public string GroupName { get; set; }

        [JsonProperty("isDeleted")]
        public bool IsDeleted { get; set; }

        [JsonProperty("ownerPrincipalName")]
        public string OwnerPrincipalName { get; set; }

        [JsonProperty("lastActivityDate")]
        public Nullable<DateTime> LastActivityDate { get; set; }

        [JsonProperty("groupType")]
        public string GroupType { get; set; }

        [JsonProperty("memberCount")]
        public Nullable<Int64> MemberCount { get; set; }

        [JsonProperty("externalMemberCount")]
        public Nullable<Int64> ExternalMemberCount { get; set; }

        [JsonProperty("exchangeReceivedEmailCount")]
        public Nullable<Int64> ExchangeReceivedEmailCount { get; set; }

        [JsonProperty("sharePointActiveFileCount")]
        public Nullable<Int64> SharePointActiveFileCount { get; set; }

        [JsonProperty("yammerPostedMessageCount")]
        public Nullable<Int64> YammerPostedMessageCount { get; set; }

        [JsonProperty("yammerReadMessageCount")]
        public Nullable<Int64> YammerReadMessageCount { get; set; }

        [JsonProperty("yammerLikedMessageCount")]
        public Nullable<Int64> YammerLikedMessageCount { get; set; }

        [JsonProperty("exchangeMailboxTotalItemCount")]
        public Nullable<Int64> ExchangeMailboxTotalItemCount { get; set; }

        [JsonProperty("exchangeMailboxStorageUsedInBytes")]
        public Nullable<Int64> ExchangeMailboxStorageUsed_Byte { get; set; }

        [JsonProperty("sharePointTotalFileCount")]
        public Nullable<Int64> SharePointTotalFileCount { get; set; }

        [JsonProperty("sharePointSiteStorageUsedInBytes")]
        public Nullable<Int64> SharePointSiteStorageUsed_Byte { get; set; }

        [JsonProperty("reportPeriod")]
        public string ReportPeriod { get; set; }
    }
}
