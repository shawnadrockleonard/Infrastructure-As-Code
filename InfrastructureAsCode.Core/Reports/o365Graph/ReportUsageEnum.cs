using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Reports.o365Graph
{
    /// <summary>
    /// Aggregate Enum for Report Types
    /// </summary>
    public enum ReportUsageEnum
    {
        /// <summary>
        /// default
        /// </summary>
        NONE = 0,

        /// <summary>
        /// Office 365 Activation / Services
        /// </summary>
        Office365 = 1,

        /// <summary>
        /// SharePoint
        /// </summary>
        SharePoint = 2,

        /// <summary>
        /// OneDrive
        /// </summary>
        OneDrive = 3,

        /// <summary>
        /// Skype for Business
        /// </summary>
        Skype = 4,

        /// <summary>
        /// Exchange Online
        /// </summary>
        Exchange = 5,

        /// <summary>
        /// Office 365 Groups
        /// </summary>
        Office365Groups = 6,

        /// <summary>
        /// MS Teams
        /// </summary>
        MSTeams = 7
    }
}
