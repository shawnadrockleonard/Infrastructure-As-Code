using InfrastructureAsCode.Core.oAuth;
using InfrastructureAsCode.Core.Reports.o365Graph.TenantReport;
using InfrastructureAsCode.Core.Reports.o365Graph.TenantReport.Mappings;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CSV = CsvHelper;
using CSVConfig = CsvHelper.Configuration;


namespace InfrastructureAsCode.Core.Reports.o365Graph
{
    public class ReportingProcessor : IReportVisitor, IDisposable
    {
        #region Properties

        public bool _disposed { get; set; }

        /// <summary>
        /// Trace Logger
        /// </summary>
        internal ITraceLogger Logger { get; private set; }

        /// <summary>
        /// Azure AD AAD v1 config
        /// </summary>
        internal IAzureADConfig AzureADConfig { get; private set; }

        internal ReportingStream ResponseReader { get; private set; }

        #endregion

        /// <summary>
        /// Supply Azure AD Credentials and associated logging utility
        /// </summary>
        /// <param name="config"></param>
        /// <param name="logger"></param>
        public ReportingProcessor(IAzureADConfig config, ITraceLogger logger)
        {
            Logger = logger;
            AzureADConfig = config;
            ResponseReader = new ReportingStream(config, logger);
        }

        /// <summary>
        /// Queries the reporting endpoint with the specified filters and interpolated classes
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <typeparam name="U"></typeparam>
        /// <param name="reportingFilters"></param>
        /// <param name="betaEndPoint"></param>
        /// <returns></returns>
        private ICollection<T> QueryBetaOrCSVMap<T, U>(QueryFilter reportingFilters, bool betaEndPoint = false)
            where T : JSONResult
            where U : CSVConfig.ClassMap
        {
            var successOrWillTry = false;
            var results = new List<T>();

            if (betaEndPoint)
            {
                // Switch to JSON Output
                try
                {
                    reportingFilters.FormattedOutput = ReportUsageFormatEnum.JSON;

                    var activityresults = ResponseReader.RetrieveData<T>(reportingFilters);
                    results.AddRange(activityresults);
                    successOrWillTry = true;
                }
                catch (Exception ex)
                {
                    Logger.LogError(ex, $"Failed for JSON Format with message {ex.Message}");
                }
            }

            if (!successOrWillTry)
            {
                // Switch to CSV Output
                reportingFilters.FormattedOutput = ReportUsageFormatEnum.Default;
                reportingFilters.BetaEndPoint = false;
                var CSVConfig = new CSVConfig.Configuration()
                {
                    Delimiter = ",",
                    HasHeaderRecord = true
                };
                CSVConfig.RegisterClassMap<U>();
                var resultscsv = new CSV.CsvReader(ResponseReader.RetrieveDataAsStream(reportingFilters), CSVConfig);
                results.AddRange(resultscsv.GetRecords<T>());
            }

            Logger.LogInformation($"Found {results.Count} while querying successOrWillTry:{successOrWillTry}");

            return results;
        }


        /// <summary>
        /// Will process the report based on the requested <typeparamref name="T"/> class
        ///     NOTE: <paramref name="betaEndPoint"/> true will default to using the JSON format; if an exception is raise it will retry with the CSV format
        ///     NOTE: <paramref name="betaEndPoint"/> false will skip JSON and use the CSV format with v1.0 endpoint
        /// </summary>
        /// <typeparam name="T">GraphAPIReport Models</typeparam>
        /// <param name="reportPeriod"></param>
        /// <param name="reportDate"></param>
        /// <param name="defaultRecordBatch"></param>
        /// <param name="betaEndPoint"></param>
        /// <returns></returns>
        public ICollection<T> ProcessReport<T>(ReportUsagePeriodEnum reportPeriod, Nullable<DateTime> reportDate, int defaultRecordBatch = 500, bool betaEndPoint = false)
            where T : JSONResult
        {
            var results = default(ICollection<T>);
            var reportingFilters = new QueryFilter(defaultRecordBatch, betaEndPoint)
            {
                O365Period = reportPeriod,
                Date = reportDate
            };

            if (typeof(T) == typeof(Office365ActiveUsersUserDetail))
            {
                reportingFilters.O365ReportType = ReportUsageTypeEnum.getOffice365ActiveUserDetail;
                if (betaEndPoint)
                {
                    Logger.LogWarning("{0} typically contains a large dataset; it is recommended to use the CSV format instead", reportingFilters.O365ReportType.ToString("f"));
                }
                results = QueryBetaOrCSVMap<T, Office365ActiveUsersUserDetailMap>(reportingFilters, betaEndPoint);
            }
            else if (typeof(T) == typeof(Office365ActiveUsersServicesUserCounts))
            {
                reportingFilters.O365ReportType = ReportUsageTypeEnum.getOffice365ServicesUserCounts;
                results = QueryBetaOrCSVMap<T, Office365ActiveUsersServicesUserCountsMap>(reportingFilters, betaEndPoint);
            }
            else if (typeof(T) == typeof(Office365GroupsActivityDetail))
            {
                reportingFilters.O365ReportType = ReportUsageTypeEnum.getOffice365GroupsActivityDetail;
                if (betaEndPoint)
                {
                    Logger.LogWarning("{0} typically contains a large dataset; it is recommended to use the CSV format instead", reportingFilters.O365ReportType.ToString("f"));
                }
                results = QueryBetaOrCSVMap<T, Office365GroupsActivityDetailMap>(reportingFilters, betaEndPoint);
            }
            else if (typeof(T) == typeof(Office365GroupsActivityCounts))
            {
                reportingFilters.O365ReportType = ReportUsageTypeEnum.getOffice365GroupsActivityCounts;
                results = QueryBetaOrCSVMap<T, Office365GroupsActivityCountsMap>(reportingFilters, betaEndPoint);
            }
            else if (typeof(T) == typeof(Office365ActiveUsersUserCounts))
            {
                reportingFilters.O365ReportType = ReportUsageTypeEnum.getOffice365ActiveUserCounts;
                results = QueryBetaOrCSVMap<T, Office365ActiveUsersUserCountsMap>(reportingFilters, betaEndPoint);
            }
            else if (typeof(T) == typeof(OneDriveActivityFileCounts))
            {
                reportingFilters.O365ReportType = ReportUsageTypeEnum.getOneDriveActivityFileCounts;
                results = QueryBetaOrCSVMap<T, OneDriveActivityFileCountsMap>(reportingFilters, betaEndPoint);
            }
            else if (typeof(T) == typeof(OneDriveActivityUserCounts))
            {
                reportingFilters.O365ReportType = ReportUsageTypeEnum.getOneDriveActivityUserCounts;
                results = QueryBetaOrCSVMap<T, OneDriveActivityUserCountsMap>(reportingFilters, betaEndPoint);
            }
            else if (typeof(T) == typeof(OneDriveActivityUserDetail))
            {
                reportingFilters.O365ReportType = ReportUsageTypeEnum.getOneDriveActivityUserDetail;
                if (betaEndPoint)
                {
                    Logger.LogWarning("{0} typically contains a large dataset; it is recommended to use the CSV format instead", reportingFilters.O365ReportType.ToString("f"));
                }
                results = QueryBetaOrCSVMap<T, OneDriveActivityUserDetailMap>(reportingFilters, betaEndPoint);
            }
            else if (typeof(T) == typeof(OneDriveUsageStorage))
            {
                reportingFilters.O365ReportType = ReportUsageTypeEnum.getOneDriveUsageStorage;
                results = QueryBetaOrCSVMap<T, OneDriveUsageStorageMap>(reportingFilters, betaEndPoint);
            }
            else if (typeof(T) == typeof(OneDriveUsageFileCounts))
            {
                reportingFilters.O365ReportType = ReportUsageTypeEnum.getOneDriveUsageFileCounts;
                results = QueryBetaOrCSVMap<T, OneDriveUsageFileCountsMap>(reportingFilters, betaEndPoint);
            }
            else if (typeof(T) == typeof(OneDriveUsageAccountCounts))
            {
                reportingFilters.O365ReportType = ReportUsageTypeEnum.getOneDriveUsageAccountCounts;
                results = QueryBetaOrCSVMap<T, OneDriveUsageAccountCountsMap>(reportingFilters, betaEndPoint);
            }
            else if (typeof(T) == typeof(OneDriveUsageAccountDetail))
            {
                reportingFilters.O365ReportType = ReportUsageTypeEnum.getOneDriveUsageAccountDetail;
                if (betaEndPoint)
                {
                    Logger.LogWarning("{0} typically contains a large dataset; it is recommended to use the CSV format instead", reportingFilters.O365ReportType.ToString("f"));
                }
                results = QueryBetaOrCSVMap<T, OneDriveUsageAccountDetailMap>(reportingFilters, betaEndPoint);
            }
            else if (typeof(T) == typeof(SharePointActivityFileCounts))
            {
                reportingFilters.O365ReportType = ReportUsageTypeEnum.getSharePointActivityFileCounts;
                results = QueryBetaOrCSVMap<T, SharePointActivityFileCountsMap>(reportingFilters, betaEndPoint);
            }
            else if (typeof(T) == typeof(SharePointActivityPages))
            {
                reportingFilters.O365ReportType = ReportUsageTypeEnum.getSharePointActivityPages;
                results = QueryBetaOrCSVMap<T, SharePointActivityPagesMap>(reportingFilters, betaEndPoint);
            }
            else if (typeof(T) == typeof(SharePointActivityUserCounts))
            {
                reportingFilters.O365ReportType = ReportUsageTypeEnum.getSharePointActivityUserCounts;
                results = QueryBetaOrCSVMap<T, SharePointActivityUserCountsMap>(reportingFilters, betaEndPoint);
            }
            else if (typeof(T) == typeof(SharePointActivityUserDetail))
            {
                reportingFilters.O365ReportType = ReportUsageTypeEnum.getSharePointActivityUserDetail;
                if (betaEndPoint)
                {
                    Logger.LogWarning("{0} typically contains a large dataset; it is recommended to use the CSV format instead", reportingFilters.O365ReportType.ToString("f"));
                }
                results = QueryBetaOrCSVMap<T, SharePointActivityUserDetailMap>(reportingFilters, betaEndPoint);
            }
            else if (typeof(T) == typeof(SharePointActivityUserCounts))
            {
                reportingFilters.O365ReportType = ReportUsageTypeEnum.getSharePointActivityUserCounts;
                results = QueryBetaOrCSVMap<T, SharePointActivityUserCountsMap>(reportingFilters, betaEndPoint);
            }
            else if (typeof(T) == typeof(SharePointSiteUsageSiteDetail))
            {
                reportingFilters.O365ReportType = ReportUsageTypeEnum.getSharePointSiteUsageDetail;
                results = QueryBetaOrCSVMap<T, SharePointSiteUsageSiteDetailMap>(reportingFilters, betaEndPoint);
            }
            else if (typeof(T) == typeof(SharePointSiteUsageFileCounts))
            {
                reportingFilters.O365ReportType = ReportUsageTypeEnum.getSharePointSiteUsageFileCounts;
                results = QueryBetaOrCSVMap<T, SharePointSiteUsageFileCountsMap>(reportingFilters, betaEndPoint);
            }
            else if (typeof(T) == typeof(SharePointSiteUsagePages))
            {
                reportingFilters.O365ReportType = ReportUsageTypeEnum.getSharePointSiteUsagePages;
                results = QueryBetaOrCSVMap<T, SharePointSiteUsagePagesMap>(reportingFilters, betaEndPoint);
            }
            else if (typeof(T) == typeof(SharePointSiteUsageSiteCounts))
            {
                reportingFilters.O365ReportType = ReportUsageTypeEnum.getSharePointSiteUsageSiteCounts;
                results = QueryBetaOrCSVMap<T, SharePointSiteUsageSiteCountsMap>(reportingFilters, betaEndPoint);
            }
            else if (typeof(T) == typeof(SharePointSiteUsageStorage))
            {
                reportingFilters.O365ReportType = ReportUsageTypeEnum.getSharePointSiteUsageStorage;
                results = QueryBetaOrCSVMap<T, SharePointSiteUsageStorageMap>(reportingFilters, betaEndPoint);
            }
            else if (typeof(T) == typeof(SkypeForBusinessActivityUserDetail))
            {
                reportingFilters.O365ReportType = ReportUsageTypeEnum.getSkypeForBusinessActivityUserDetail;
                if (betaEndPoint)
                {
                    Logger.LogWarning("{0} typically contains a large dataset; it is recommended to use the CSV format instead", reportingFilters.O365ReportType.ToString("f"));
                }
                results = QueryBetaOrCSVMap<T, SkypeForBusinessActivityUserDetailMap>(reportingFilters, betaEndPoint);
            }
            else if (typeof(T) == typeof(SkypeForBusinessActivityActivityCounts))
            {
                reportingFilters.O365ReportType = ReportUsageTypeEnum.getSkypeForBusinessActivityCounts;
                results = QueryBetaOrCSVMap<T, SkypeForBusinessActivityActivityCountsMap>(reportingFilters, betaEndPoint);
            }
            else if (typeof(T) == typeof(SkypeForBusinessActivityUserCounts))
            {
                reportingFilters.O365ReportType = ReportUsageTypeEnum.getSkypeForBusinessActivityUserCounts;
                results = QueryBetaOrCSVMap<T, SkypeForBusinessActivityUserCountsMap>(reportingFilters, betaEndPoint);
            }
            else if (typeof(T) == typeof(SkypeForBusinessDeviceUsageUserDetail))
            {
                reportingFilters.O365ReportType = ReportUsageTypeEnum.getSkypeForBusinessDeviceUsageUserDetail;
                if (betaEndPoint)
                {
                    Logger.LogWarning("{0} typically contains a large dataset; it is recommended to use the CSV format instead", reportingFilters.O365ReportType.ToString("f"));
                }
                results = QueryBetaOrCSVMap<T, SkypeForBusinessDeviceUsageUserDetailMap>(reportingFilters, betaEndPoint);
            }
            else if (typeof(T) == typeof(SkypeForBusinessDeviceUsageDistributionUserCounts))
            {
                reportingFilters.O365ReportType = ReportUsageTypeEnum.getSkypeForBusinessDeviceUsageDistributionUserCounts;
                results = QueryBetaOrCSVMap<T, SkypeForBusinessDeviceUsageDistributionUserCountsMap>(reportingFilters, betaEndPoint);
            }
            else if (typeof(T) == typeof(SkypeForBusinessDeviceUsageUserCounts))
            {
                reportingFilters.O365ReportType = ReportUsageTypeEnum.getSkypeForBusinessDeviceUsageUserCounts;
                results = QueryBetaOrCSVMap<T, SkypeForBusinessDeviceUsageUserCountsMap>(reportingFilters, betaEndPoint);
            }
            else if (typeof(T) == typeof(SkypeForBusinessOrganizerActivityCounts))
            {
                reportingFilters.O365ReportType = ReportUsageTypeEnum.getSkypeForBusinessOrganizerActivityCounts;
                results = QueryBetaOrCSVMap<T, SkypeForBusinessOrganizerActivityCountsMap>(reportingFilters, betaEndPoint);
            }
            else if (typeof(T) == typeof(SkypeForBusinessOrganizerActivityUserCounts))
            {
                reportingFilters.O365ReportType = ReportUsageTypeEnum.getSkypeForBusinessOrganizerActivityUserCounts;
                results = QueryBetaOrCSVMap<T, SkypeForBusinessOrganizerActivityUserCountsMap>(reportingFilters, betaEndPoint);
            }
            else if (typeof(T) == typeof(SkypeForBusinessOrganizerActivityMinuteCounts))
            {
                reportingFilters.O365ReportType = ReportUsageTypeEnum.getSkypeForBusinessOrganizerActivityMinuteCounts;
                results = QueryBetaOrCSVMap<T, SkypeForBusinessOrganizerActivityMinuteCountsMap>(reportingFilters, betaEndPoint);
            }
            else if (typeof(T) == typeof(SkypeForBusinessParticipantActivityCounts))
            {
                reportingFilters.O365ReportType = ReportUsageTypeEnum.getSkypeForBusinessParticipantActivityCounts;
                results = QueryBetaOrCSVMap<T, SkypeForBusinessParticipantActivityCountsMap>(reportingFilters, betaEndPoint);
            }
            else if (typeof(T) == typeof(SkypeForBusinessParticipantActivityUserCounts))
            {
                reportingFilters.O365ReportType = ReportUsageTypeEnum.getSkypeForBusinessParticipantActivityUserCounts;
                results = QueryBetaOrCSVMap<T, SkypeForBusinessParticipantActivityUserCountsMap>(reportingFilters, betaEndPoint);
            }
            else if (typeof(T) == typeof(SkypeForBusinessParticipantActivityMinuteCounts))
            {
                reportingFilters.O365ReportType = ReportUsageTypeEnum.getSkypeForBusinessParticipantActivityMinuteCounts;
                results = QueryBetaOrCSVMap<T, SkypeForBusinessParticipantActivityMinuteCountsMap>(reportingFilters, betaEndPoint);
            }
            else if (typeof(T) == typeof(SkypeForBusinessPeerToPeerActivityCounts))
            {
                reportingFilters.O365ReportType = ReportUsageTypeEnum.getSkypeForBusinessPeerToPeerActivityCounts;
                results = QueryBetaOrCSVMap<T, SkypeForBusinessPeerToPeerActivityCountsMap>(reportingFilters, betaEndPoint);
            }
            else if (typeof(T) == typeof(SkypeForBusinessPeerToPeerActivityUserCounts))
            {
                reportingFilters.O365ReportType = ReportUsageTypeEnum.getSkypeForBusinessPeerToPeerActivityUserCounts;
                results = QueryBetaOrCSVMap<T, SkypeForBusinessPeerToPeerActivityUserCountsMap>(reportingFilters, betaEndPoint);
            }
            else if (typeof(T) == typeof(SkypeForBusinessPeerToPeerActivityMinuteCounts))
            {
                reportingFilters.O365ReportType = ReportUsageTypeEnum.getSkypeForBusinessPeerToPeerActivityMinuteCounts;
                results = QueryBetaOrCSVMap<T, SkypeForBusinessPeerToPeerActivityMinuteCountsMap>(reportingFilters, betaEndPoint);
            }

            return results;
        }

        /// <summary>
        /// Queries the Graph API and returns the emitted string output
        /// </summary>
        /// <param name="ReportType"></param>
        /// <param name="reportPeriod"></param>
        /// <param name="reportDate"></param>
        /// <param name="defaultRecordBatch">(OPTIONAL) will default to 500</param>
        /// <param name="betaEndPoint">(OPTIONAL) will default to false</param>
        /// <returns></returns>
        public string ProcessReport(ReportUsageTypeEnum ReportType, ReportUsagePeriodEnum reportPeriod, Nullable<DateTime> reportDate, int defaultRecordBatch = 500, bool betaEndPoint = false)
        {
            var serviceQuery = new QueryFilter(defaultRecordBatch, betaEndPoint)
            {
                O365ReportType = ReportType,
                O365Period = reportPeriod,
                Date = reportDate
            };

            var response = ResponseReader.RetrieveData(serviceQuery);
            return response;
        }


        /// <summary>
        /// Dispose of the class
        /// </summary>
        /// <param name="disposing"></param>
        protected virtual void Dispose(bool disposing)
        {
            if (disposing
                && !_disposed)
            {
            }
            _disposed = true;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
    }
}
