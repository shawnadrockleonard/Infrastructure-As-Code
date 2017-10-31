using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Reports.o365Graph
{
    /// <summary>
    /// Represents the O365 Usage API EndPoints
    /// </summary>
    public enum ReportUsageTypeEnum
    {
        NONE,
        EmailActivity,
        getEmailActivityUserDetail,
        getEmailActivityCounts,
        getEmailActivityUserCounts,
        getEmailAppUsageUserDetail,
        getEmailAppUsageAppsUserCounts,
        getEmailAppUsageUserCounts,
        getEmailAppUsageVersionsUserCounts,
        getMailboxUsageDetail,
        getMailboxUsageMailboxCounts,
        getMailboxUsageQuotaMailboxStatusCounts,
        getMailboxUsageStorage,
        getOffice365GroupsActivityDetail,
        getOffice365GroupsActivityCounts,
        getOffice365GroupsActivityGroupCounts,
        getOffice365GroupsActivityStorage,
        getOffice365GroupsActivityFileCounts,
        getOneDriveActivityUserDetail,
        getOneDriveActivityUserCounts,
        getOneDriveActivityFileCounts,
        getOneDriveUsageAccountDetail,
        getOneDriveUsageAccountCounts,
        getOneDriveUsageFileCounts,
        getOneDriveUsageStorage,
        getSharePointActivityUserDetail,
        getSharePointActivityFileCounts,
        getSharePointActivityUserCounts,
        getSharePointActivityPages,
        getSharePointSiteUsageDetail,
        getSharePointSiteUsageFileCounts,
        getSharePointSiteUsageSiteCounts,
        getSharePointSiteUsageStorage,
        getSharePointSiteUsagePages,
        getSkypeForBusinessActivityUserDetail,
        getSkypeForBusinessActivityCounts,
        getSkypeForBusinessActivityUserCounts,
        getSkypeForBusinessDeviceUsageUserDetail,
        getSkypeForBusinessDeviceUsageDistributionUserCounts,
        getSkypeForBusinessDeviceUsageUserCounts,
        getSkypeForBusinessOrganizerActivityCounts,
        getSkypeForBusinessOrganizerActivityUserCounts,
        getSkypeForBusinessOrganizerActivityMinuteCounts,
        getSkypeForBusinessParticipantActivityCounts,
        getSkypeForBusinessParticipantActivityUserCounts,
        getSkypeForBusinessParticipantActivityMinuteCounts,
        getSkypeForBusinessPeerToPeerActivityCounts,
        getSkypeForBusinessPeerToPeerActivityUserCounts,
        getSkypeForBusinessPeerToPeerActivityMinuteCounts,
        getOffice365ActivationsUserDetail,
        getOffice365ActivationCounts,
        getOffice365ActivationsUserCounts,
        getOffice365ActiveUserDetail,
        getOffice365ActiveUserCounts,
        getOffice365ServicesUserCounts,
        getYammerActivityUserDetail,
        getYammerActivityCounts,
        getYammerActivityUserCounts,
        getYammerDeviceUsageUserDetail,
        getYammerDeviceUsageDistributionUserCounts,
        getYammerDeviceUsageUserCounts,
        getYammerGroupsActivityDetail,
        getYammerGroupsActivityGroupCounts,
        getYammerGroupsActivityCounts
    }
}
