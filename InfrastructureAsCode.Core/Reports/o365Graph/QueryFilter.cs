using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Reports.o365Graph
{
    internal class QueryFilter
    {
        #region Properties 

        /// <summary>
        /// Represents the Graph API endpoints
        /// </summary>
        /// <remarks>Of note this is a BETA inpoint as these APIs are in Preview</remarks>
        private static readonly string DefaultServiceEndpointUrl = "https://graph.microsoft.com/{0}/reports/{1}{2}";

        /// <summary>
        /// Represents Static URI endpoints for reports
        /// </summary>
        internal ReportUsageTypeEnum O365ReportType { get; set; }

        /// <summary>
        /// Period associated with the endpoint D7, D30, D90
        /// </summary>
        internal Nullable<ReportUsagePeriodEnum> O365Period { get; set; }

        /// <summary>
        /// Formats the response
        /// </summary>
        /// <remarks>Default CSV</remarks>
        internal ReportUsageFormatEnum FormattedOutput { get; set; }

        /// <summary>
        /// A Date filter to be applied to the set of URI OData filters
        /// </summary>
        internal Nullable<DateTime> Date { get; set; }

        /// <summary>
        /// Represents the $top total number of records to return in the API call
        /// </summary>
        internal int RecordBatchCount { get; set; }

        /// <summary>
        /// Should we default to the v1.0 endpoint or use the beta endpoint
        /// </summary>
        internal bool BetaEndPoint { get; set; }

        #endregion

        /// <summary>
        /// Initialize the collection with defaults
        /// </summary>
        /// <param name="defaultRecordBatch">Defaults 100</param>
        /// <param name="betaEndpoint">(optional) should we consume the beta endpoint or v1.0</param>
        internal QueryFilter(int defaultRecordBatch = 100, bool betaEndpoint = false)
        {
            RecordBatchCount = defaultRecordBatch;
            BetaEndPoint = betaEndpoint;
        }

        /// <summary>
        /// #Build the request URL and invoke
        ///     Sample: OneDriveActivity(view='Detail',period='D7')/content
        /// </summary>
        /// <returns></returns>
        internal Uri ToUrl()
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
        internal Uri ToUrl(string graphUrl)
        {
            var parameterset = string.Empty;

            // If period is specified then add that to the parameters except when not is supported
            var doesNotSupportPeriod = new ReportUsageTypeEnum[]
            {
                ReportUsageTypeEnum.getOffice365ActivationsUserDetail,
                ReportUsageTypeEnum.getOffice365ActivationCounts,
                ReportUsageTypeEnum.getOffice365ActivationsUserCounts
            };
            if (O365Period.HasValue && !doesNotSupportPeriod.Any(a => a == O365ReportType))
            {
                parameterset = string.Format("period='{0}'", O365Period.Value.ToString("f"));
            }

            // If the date is specified then add that to the parameters if it is supported
            var doesSupportDate = new ReportUsageTypeEnum[]
            {
                ReportUsageTypeEnum.getOffice365ActiveUserDetail,
                ReportUsageTypeEnum.getOffice365GroupsActivityDetail,
                ReportUsageTypeEnum.getOneDriveActivityUserDetail,
                ReportUsageTypeEnum.getOneDriveUsageAccountDetail,
                ReportUsageTypeEnum.getSharePointActivityUserDetail,
                ReportUsageTypeEnum.getSharePointSiteUsageDetail,
                ReportUsageTypeEnum.getSkypeForBusinessActivityUserDetail,
                ReportUsageTypeEnum.getSkypeForBusinessDeviceUsageUserDetail
            };
            if (Date.HasValue && doesSupportDate.Any(a => a == O365ReportType))
            {
                parameterset = string.Format("date={0}", Date.Value.ToString("yyyy-MM-dd"));
            }


            if (FormattedOutput == ReportUsageFormatEnum.JSON)
            {
                var supportsTop = new ReportUsageTypeEnum[]
                {
                    ReportUsageTypeEnum.getOffice365ActiveUserDetail,
                    ReportUsageTypeEnum.getOffice365GroupsActivityDetail,
                    ReportUsageTypeEnum.getOneDriveActivityUserDetail,
                    ReportUsageTypeEnum.getOneDriveUsageAccountDetail,
                    ReportUsageTypeEnum.getSharePointSiteUsageDetail,
                    ReportUsageTypeEnum.getSkypeForBusinessActivityUserDetail,
                    ReportUsageTypeEnum.getSkypeForBusinessDeviceUsageUserDetail
                };
                var topIsDirty = false;
                if (supportsTop.Any(a => a == O365ReportType))
                {
                    topIsDirty = true;
                    graphUrl += "?$top=" + RecordBatchCount;
                }
                graphUrl += (topIsDirty ? "&" : "?") + "$format=application/json";
            }


            // JSON Format not supported by V1.0 endpoint
            var versionuri = (BetaEndPoint || (FormattedOutput == ReportUsageFormatEnum.JSON) ? "beta" : "v1.0");
            // If parameter is specified enable parenthesis
            if (!string.IsNullOrEmpty(parameterset))
            {
                parameterset = $"({parameterset})";
            }

            var uri = new Uri(string.Format(graphUrl, versionuri, O365ReportType, parameterset));
            return uri;
        }
    }
}
