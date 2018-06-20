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
     * Report Refresh Date,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,Report Period
     * */
    class SkypeForBusinessActivityUserDetailMap : ClassMap<SkypeForBusinessActivityUserDetail>
    {
        public SkypeForBusinessActivityUserDetailMap()
        {
            Map(m => m.ReportRefreshDate).Name("Report Refresh Date").Index(0).Default(default(DateTime));
            Map(m => m.UPN).Name("User Principal Name").Index(1).Default(string.Empty);
            Map(m => m.Deleted).Name("Is Deleted").Index(2).Default("false");
            Map(m => m.DeletedDate).Name("Deleted Date").Index(3).Default(default(Nullable<DateTime>));
            Map(m => m.LastActivityDate).Name("Last Activity Date").Index(4).Default(default(Nullable<DateTime>));
            Map(m => m.TotalPeerToPeerSessionCount).Name("Total Peer-to-peer Session Count").Index(5).Default(0);
            Map(m => m.TotalOrganizedConferenceCount).Name("Total Organized Conference Count").Index(6).Default(0);
            Map(m => m.TotalParticipatedConferenceCount).Name("Total Participated Conference Count").Index(7).Default(0);
            Map(m => m.PeerToPeerLastActivityDate).Name("Peer-to-peer Last Activity Date").Index(8).Default(default(Nullable<DateTime>));
            Map(m => m.OrganizedConferenceLastActivityDate).Name("Organized Conference Last Activity Date").Index(9).Default(default(Nullable<DateTime>));
            Map(m => m.ParticipatedConferenceLastActivityDate).Name("Participated Conference Last Activity Date").Index(10).Default(default(Nullable<DateTime>));
            Map(m => m.PeerToPeerIMCount).Name("Peer-to-peer IM Count").Index(11).Default(0);
            Map(m => m.PeerToPeerAudioCount).Name("Peer-to-peer Audio Count").Index(12).Default(0);
            Map(m => m.PeerToPeerAudioMinutes).Name("Peer-to-peer Audio Minutes").Index(13).Default(0);
            Map(m => m.PeerToPeerVideoCount).Name("Peer-to-peer Video Count").Index(14).Default(0);
            Map(m => m.PeerToPeerVideoMinutes).Name("Peer-to-peer Video Minutes").Index(15).Default(0);
            Map(m => m.PeerToPeerAppSharingCount).Name("Peer-to-peer App Sharing Count").Index(16).Default(0);
            Map(m => m.PeerToPeerFileTransferCount).Name("Peer-to-peer File Transfer Count").Index(17).Default(0);
            Map(m => m.OrganizedConferenceIMCount).Name("Organized Conference IM Count").Index(18).Default(0);
            Map(m => m.OrganizedConferenceAudioVideoCount).Name("Organized Conference Audio/Video Count").Index(19).Default(0);
            Map(m => m.OrganizedConferenceAudioVideoMinutes).Name("Organized Conference Audio/Video Minutes").Index(20).Default(0);
            Map(m => m.OrganizedConferenceAppSharingCount).Name("Organized Conference App Sharing Count").Index(21).Default(0);
            Map(m => m.OrganizedConferenceWebCount).Name("Organized Conference Web Count").Index(22).Default(0);
            Map(m => m.OrganizedConferenceDialInOut3rdPartyCount).Name("Organized Conference Dial-in/out 3rd Party Count").Index(23).Default(0);
            Map(m => m.OrganizedConferenceCloudDialInOutMicrosoftCount).Name("Organized Conference Dial-in/out Microsoft Count").Index(24).Default(0);
            Map(m => m.OrganizedConferenceCloudDialInMicrosoftMinutes).Name("Organized Conference Dial-in Microsoft Minutes").Index(25).Default(0);
            Map(m => m.OrganizedConferenceCloudDialOutMicrosoftMinutes).Name("Organized Conference Dial-out Microsoft Minutes").Index(26).Default(0);
            Map(m => m.ParticipatedConferenceIMCount).Name("Participated Conference IM Count").Index(27).Default(0);
            Map(m => m.ParticipatedConferenceAudioVideoCount).Name("Participated Conference Audio/Video Count").Index(28).Default(0);
            Map(m => m.ParticipatedConferenceAudioVideoMinutes).Name("Participated Conference Audio/Video Minutes").Index(29).Default(0);
            Map(m => m.ParticipatedConferenceAppSharingCount).Name("Participated Conference App Sharing Count").Index(30).Default(0);
            Map(m => m.ParticipatedConferenceWebCount).Name("Participated Conference Web Count").Index(31).Default(0);
            Map(m => m.ParticipatedConferenceDialInOut3rdPartyCount).Name("Participated Conference Dial-in/out 3rd Party Count").Index(32).Default(0);
            Map(m => m.ProductsAssignedCSV).Name("Assigned Products").Index(33).Default(string.Empty);
            Map(m => m.ReportPeriod).Name("Report Period").Index(34).Default(0);
        }
    }


    /// <summary>
    /// Get details about Skype for Business activity by user.
    /// </summary>
    public class SkypeForBusinessActivityUserDetail : JSONODataBase
    {
        [JsonProperty("reportRefreshDate")]
        public DateTime ReportRefreshDate { get; set; }

        [JsonProperty("userPrincipalName")]
        public string UPN { get; set; }

        [JsonProperty("isDeleted")]
        public string Deleted { get; set; }

        [JsonProperty("deletedDate")]
        public Nullable<DateTime> DeletedDate { get; set; }

        [JsonProperty("lastActivityDate")]
        public Nullable<DateTime> LastActivityDate { get; set; }

        [JsonProperty("totalPeerToPeerSessionCount")]
        public Nullable<Int64> TotalPeerToPeerSessionCount { get; set; }

        [JsonProperty("totalOrganizedConferenceCount")]
        public Nullable<Int64> TotalOrganizedConferenceCount { get; set; }

        [JsonProperty("totalParticipatedConferenceCount")]
        public Nullable<Int64> TotalParticipatedConferenceCount { get; set; }

        [JsonProperty("peerToPeerLastActivityDate")]
        public Nullable<DateTime> PeerToPeerLastActivityDate { get; set; }

        [JsonProperty("organizedConferenceLastActivityDate")]
        public Nullable<DateTime> OrganizedConferenceLastActivityDate { get; set; }

        [JsonProperty("participatedConferenceLastActivityDate")]
        public Nullable<DateTime> ParticipatedConferenceLastActivityDate { get; set; }

        [JsonProperty("peerToPeerIMCount")]
        public Nullable<Int64> PeerToPeerIMCount { get; set; }

        [JsonProperty("peerToPeerAudioCount")]
        public Nullable<Int64> PeerToPeerAudioCount { get; set; }

        [JsonProperty("peerToPeerAudioMinutes")]
        public Nullable<Int64> PeerToPeerAudioMinutes { get; set; }

        [JsonProperty("peerToPeerVideoCount")]
        public Nullable<Int64> PeerToPeerVideoCount { get; set; }

        [JsonProperty("peerToPeerVideoMinutes")]
        public Nullable<Int64> PeerToPeerVideoMinutes { get; set; }

        [JsonProperty("peerToPeerAppSharingCount")]
        public Nullable<Int64> PeerToPeerAppSharingCount { get; set; }

        [JsonProperty("peerToPeerFileTransferCount")]
        public Nullable<Int64> PeerToPeerFileTransferCount { get; set; }

        [JsonProperty("organizedConferenceIMCount")]
        public Nullable<Int64> OrganizedConferenceIMCount { get; set; }

        [JsonProperty("organizedConferenceAudioVideoCount")]
        public Nullable<Int64> OrganizedConferenceAudioVideoCount { get; set; }

        [JsonProperty("organizedConferenceAudioVideoMinutes")]
        public Nullable<Int64> OrganizedConferenceAudioVideoMinutes { get; set; }

        [JsonProperty("organizedConferenceAppSharingCount")]
        public Nullable<Int64> OrganizedConferenceAppSharingCount { get; set; }

        [JsonProperty("organizedConferenceWebCount")]
        public Nullable<Int64> OrganizedConferenceWebCount { get; set; }

        [JsonProperty("organizedConferenceDialInOut3rdPartyCount")]
        public Nullable<Int64> OrganizedConferenceDialInOut3rdPartyCount { get; set; }

        [JsonProperty("organizedConferenceCloudDialInOutMicrosoftCount")]
        public Nullable<Int64> OrganizedConferenceCloudDialInOutMicrosoftCount { get; set; }

        [JsonProperty("organizedConferenceCloudDialInMicrosoftMinutes")]
        public Nullable<Int64> OrganizedConferenceCloudDialInMicrosoftMinutes { get; set; }

        [JsonProperty("organizedConferenceCloudDialOutMicrosoftMinutes")]
        public Nullable<Int64> OrganizedConferenceCloudDialOutMicrosoftMinutes { get; set; }

        [JsonProperty("participatedConferenceIMCount")]
        public Nullable<Int64> ParticipatedConferenceIMCount { get; set; }

        [JsonProperty("participatedConferenceAudioVideoCount")]
        public Nullable<Int64> ParticipatedConferenceAudioVideoCount { get; set; }

        [JsonProperty("participatedConferenceAudioVideoMinutes")]
        public Nullable<Int64> ParticipatedConferenceAudioVideoMinutes { get; set; }

        [JsonProperty("participatedConferenceAppSharingCount")]
        public Nullable<Int64> ParticipatedConferenceAppSharingCount { get; set; }

        [JsonProperty("participatedConferenceWebCount")]
        public Nullable<Int64> ParticipatedConferenceWebCount { get; set; }


        [JsonProperty("participatedConferenceDialInOut3rdPartyCount")]
        public Nullable<Int64> ParticipatedConferenceDialInOut3rdPartyCount { get; set; }

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