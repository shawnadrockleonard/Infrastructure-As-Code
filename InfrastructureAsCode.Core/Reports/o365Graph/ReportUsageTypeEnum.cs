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
        EmailAppUsage,
        EmailActivity,
        MailboxUsage,
        Office365ActiveUsers,
        Office365GroupsActivity,
        Office365Activations,
        OneDriveUsage,
        OneDriveActivity,
        SfbDeviceUsage,
        SfbOrganizerActivity,
        SfbP2PActivity,
        SfbParticipantActivity,
        SfbActivity,
        SharePointSiteUsage,
        SharePointActivity,
        YammerDeviceUsage,
        YammerActivity
    }
}
