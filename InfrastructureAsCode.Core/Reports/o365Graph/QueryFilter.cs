using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Reports.o365Graph
{
    public class QueryFilter
    {
        #region Properties 

        /// <summary>
        /// Represents the Graph API endpoints
        /// </summary>
        /// <remarks>Of note this is a BETA inpoint as these APIs are in Preview</remarks>
        public static string DefaultServiceEndpointUrl = "https://graph.microsoft.com/beta/reports/{0}({1})";

        public ReportUsageTypeEnum O365ReportType { get; set; }

        public Nullable<ReportUsagePeriodEnum> O365Period { get; set; }

        /// <summary>
        /// Formats the response
        /// </summary>
        /// <remarks>Default CSV</remarks>
        public ReportUsageFormatEnum FormattedOutput { get; set; }

        /// <summary>
        /// A Date filter to be applied to the set of URI OData filters
        /// </summary>
        public Nullable<DateTime> Date { get; set; }

        /// <summary>
        /// Represents the $top total number of records to return in the API call
        /// </summary>
        public int RecordBatchCount { get; set; }

        #endregion

        /// <summary>
        /// Initialize the collection with defaults
        /// </summary>
        public QueryFilter()
        {
            RecordBatchCount = 100;
        }

        /// <summary>
        /// #Build the request URL and invoke
        ///     Sample: OneDriveActivity(view='Detail',period='D7')/content
        /// </summary>
        /// <returns></returns>
        public Uri ToUrl()
        {
            var uri = ToUrl(DefaultServiceEndpointUrl);
            return uri;
        }

        /// <summary>
        /// #Build the request URL and invoke
        ///     Sample: OneDriveActivity(view='Detail',period='D7')/content
        /// </summary>
        /// <param name="graphUrl">Represents the Graph URL for Usage Reporting which should have two parameters {0}{1}</param>
        /// <returns></returns>
        public Uri ToUrl(string graphUrl)
        {
            var str = string.Empty;

            // We always have a view to start with that
            var parameterset = string.Empty;

            // If period is specified then add that to the parameters unless it is not supported
            var activations = new ReportUsageTypeEnum[]
            {
                ReportUsageTypeEnum.getOffice365ActivationsUserDetail,
                ReportUsageTypeEnum.getOffice365ActivationCounts,
                ReportUsageTypeEnum.getOffice365ActivationsUserCounts
            };
            if (!Date.HasValue && O365Period.HasValue && !activations.Any(a => a == O365ReportType))
            {
                str = string.Format("period='{0}',", O365Period.Value.ToString("f"));
                parameterset += str;
            }

            // If the date is specified then add that to the parameters unless it is not supported
            var mailboxes = new ReportUsageTypeEnum[]
            {
                ReportUsageTypeEnum.getMailboxUsageDetail,
                ReportUsageTypeEnum.getMailboxUsageMailboxCounts,
                ReportUsageTypeEnum.getMailboxUsageQuotaMailboxStatusCounts,
                ReportUsageTypeEnum.getMailboxUsageStorage
            };
            var skypeactivities = new ReportUsageTypeEnum[]
            {
                ReportUsageTypeEnum.getSkypeForBusinessOrganizerActivityCounts,
                ReportUsageTypeEnum.getSkypeForBusinessOrganizerActivityUserCounts,
                ReportUsageTypeEnum.getSkypeForBusinessOrganizerActivityMinuteCounts
            };
            if (Date.HasValue
                && !(mailboxes.Any(a => a == O365ReportType)
                  || activations.Any(a => a == O365ReportType)
                  || skypeactivities.Any(a => a == O365ReportType)))
            {
                str = string.Format("date={0}", Date.Value.ToString("yyyy-MM-dd"));
                parameterset += str;
            }

            if (FormattedOutput == ReportUsageFormatEnum.JSON)
            {
                var supportsTop = new ReportUsageTypeEnum[]
                {
                    ReportUsageTypeEnum.getOneDriveActivityUserDetail,
                    ReportUsageTypeEnum.getOneDriveUsageAccountDetail
                };
                var topIsDirty = false;
                if (supportsTop.Any(a => a == O365ReportType))
                {
                    topIsDirty = true;
                    graphUrl += "?$top=" + RecordBatchCount;
                }
                graphUrl += (topIsDirty ? "&" : "?") + "$format=application/json";
            }

            // #Trim a trailing comma off the ParameterSet if needed
            parameterset = parameterset.TrimEnd(new char[] { ',' });

            var uri = new Uri(string.Format(graphUrl, O365ReportType, parameterset));
            return uri;
        }
    }
}
