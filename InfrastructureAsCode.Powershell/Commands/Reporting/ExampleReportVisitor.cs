using InfrastructureAsCode.Core.Reports;
using InfrastructureAsCode.Core.Reports.o365Graph;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CsvHelper;
using CsvHelper.Configuration;
using InfrastructureAsCode.Powershell.Commands.Reporting.Usage;
using InfrastructureAsCode.Core.Reports.o365Graph.TenantReport;

namespace InfrastructureAsCode.Powershell.Commands.Reporting
{
    public class ExampleReportVisitor : ReportVisitor, IDisposable
    {
        #region Properties

        internal Configuration CSVConfig { get; private set; }

        public bool _disposed { get; set; }

        #endregion

        public ExampleReportVisitor() { }

        public ExampleReportVisitor(ITraceLogger logger)
        {
            Logger = logger;
        }

        public ExampleReportVisitor(Configuration _csvconfig, ITraceLogger _logger) : this(_logger)
        {
            CSVConfig = _csvconfig;
        }

        public override void ProcessReport(ReportingStream responseReader, QueryFilter reportingFilters)
        {

            if (reportingFilters.O365ReportType == ReportUsageTypeEnum.getOneDriveActivityFileCounts)
            {
                // Switch to JSON output
                reportingFilters.FormattedOutput = ReportUsageFormatEnum.JSON;
                var activityresults = responseReader.RetrieveData<OneDriveActivityFileCounts>(reportingFilters);
                var results = activityresults.value.ToList();
                var rowprocessed = results.Count();

                // Switch to CSV Output
                reportingFilters.FormattedOutput = ReportUsageFormatEnum.Default;
                CSVConfig.RegisterClassMap<OneDriveActivityFileCountsMap>();
                var resultscsv = new CsvReader(responseReader.RetrieveDataAsStream(reportingFilters), CSVConfig);
                var resultsrec = resultscsv.GetRecords<OneDriveActivityFileCounts>().ToList();
            }
            else if (reportingFilters.O365ReportType == ReportUsageTypeEnum.getOneDriveActivityUserCounts)
            {
                // Switch to JSON output
                reportingFilters.FormattedOutput = ReportUsageFormatEnum.JSON;
                var activityresults = responseReader.RetrieveData<OneDriveActivityUserCounts>(reportingFilters);
                var results = activityresults.value.ToList();
                var rowprocessed = results.Count();

                // Switch to CSV Output
                reportingFilters.FormattedOutput = ReportUsageFormatEnum.Default;
                CSVConfig.RegisterClassMap<OneDriveActivityUserCountsMap>();
                var resultscsv = new CsvReader(responseReader.RetrieveDataAsStream(reportingFilters), CSVConfig);
                var resultsrec = resultscsv.GetRecords<OneDriveActivityUserCounts>().ToList();
            }
            else if (reportingFilters.O365ReportType == ReportUsageTypeEnum.getOneDriveActivityUserDetail)
            {
                // Switch to JSON output
                reportingFilters.FormattedOutput = ReportUsageFormatEnum.JSON;
                var activityresults = responseReader.RetrieveData<OneDriveActivityDetail>(reportingFilters);
                var results = activityresults.value.ToList();
                var rowprocessed = results.Count();

                // Switch to CSV Output
                reportingFilters.FormattedOutput = ReportUsageFormatEnum.Default;
                CSVConfig.RegisterClassMap<OneDriveActivityDetailMap>();
                var resultscsv = new CsvReader(responseReader.RetrieveDataAsStream(reportingFilters), CSVConfig);
                var resultsrec = resultscsv.GetRecords<OneDriveActivityDetail>().ToList();
            }
            else if (reportingFilters.O365ReportType == ReportUsageTypeEnum.getOneDriveUsageStorage)
            {
                // Switch to JSON output
                reportingFilters.FormattedOutput = ReportUsageFormatEnum.JSON;
                var activityresults = responseReader.RetrieveData<OneDriveUsageStorage>(reportingFilters);
                var results = activityresults.value.ToList();
                var rowprocessed = results.Count();

                // Switch to CSV Output
                reportingFilters.FormattedOutput = ReportUsageFormatEnum.Default;
                CSVConfig.RegisterClassMap<OneDriveUsageStorageMap>();
                var resultscsv = new CsvReader(responseReader.RetrieveDataAsStream(reportingFilters), CSVConfig);
                var resultsrec = resultscsv.GetRecords<OneDriveUsageStorage>().ToList();
            }
            else if (reportingFilters.O365ReportType == ReportUsageTypeEnum.getOneDriveUsageFileCounts)
            {
                // Switch to JSON output
                reportingFilters.FormattedOutput = ReportUsageFormatEnum.JSON;
                var activityresults = responseReader.RetrieveData<OneDriveUsageFileCounts>(reportingFilters);
                var results = activityresults.value.ToList();
                var rowprocessed = results.Count();

                // Switch to CSV Output
                reportingFilters.FormattedOutput = ReportUsageFormatEnum.Default;
                CSVConfig.RegisterClassMap<OneDriveUsageFileCountsMap>();
                var resultscsv = new CsvReader(responseReader.RetrieveDataAsStream(reportingFilters), CSVConfig);
                var resultsrec = resultscsv.GetRecords<OneDriveUsageFileCounts>().ToList();
            }
            else if (reportingFilters.O365ReportType == ReportUsageTypeEnum.getOneDriveUsageAccountCounts)
            {
                // Switch to JSON output
                reportingFilters.FormattedOutput = ReportUsageFormatEnum.JSON;
                var activityresults = responseReader.RetrieveData<OneDriveUsageAccountCounts>(reportingFilters);
                var results = activityresults.value.ToList();
                var rowprocessed = results.Count();

                // Switch to CSV Output
                reportingFilters.FormattedOutput = ReportUsageFormatEnum.Default;
                CSVConfig.RegisterClassMap<OneDriveUsageAccountCountsMap>();
                var resultscsv = new CsvReader(responseReader.RetrieveDataAsStream(reportingFilters), CSVConfig);
                var resultsrec = resultscsv.GetRecords<OneDriveUsageAccountCounts>().ToList();
            }
            else if (reportingFilters.O365ReportType == ReportUsageTypeEnum.getOneDriveUsageAccountDetail)
            {
                // Switch to JSON output
                reportingFilters.FormattedOutput = ReportUsageFormatEnum.JSON;
                var activityresults = responseReader.RetrieveData<OneDriveUsageAccountDetail>(reportingFilters);
                var results = activityresults.value.ToList();
                var rowprocessed = results.Count();

                // Switch to CSV Output
                reportingFilters.FormattedOutput = ReportUsageFormatEnum.Default;
                CSVConfig.RegisterClassMap<OneDriveUsageAccountDetailMap>();
                var resultscsv = new CsvReader(responseReader.RetrieveDataAsStream(reportingFilters), CSVConfig);
                var resultsrec = resultscsv.GetRecords<OneDriveUsageAccountDetail>().ToList();
            }
            else if (reportingFilters.O365ReportType == ReportUsageTypeEnum.getOffice365ActiveUserDetail)
            {
                // Switch to JSON output
                reportingFilters.FormattedOutput = ReportUsageFormatEnum.JSON;
                var activityresults = responseReader.RetrieveData<Office365ActiveUsersDetails>(reportingFilters);
                var results = activityresults.value.ToList();
                var rowprocessed = results.Count();

                // Switch to CSV Output
                reportingFilters.FormattedOutput = ReportUsageFormatEnum.Default;
                CSVConfig.RegisterClassMap<Office365ActiveUsersDetailsMap>();
                var resultscsv = new CsvReader(responseReader.RetrieveDataAsStream(reportingFilters), CSVConfig);
                var resultsrec = resultscsv.GetRecords<Office365ActiveUsersDetails>().ToList();
            }
            else if (reportingFilters.O365ReportType == ReportUsageTypeEnum.getOffice365ServicesUserCounts)
            {
                // Switch to JSON output
                reportingFilters.FormattedOutput = ReportUsageFormatEnum.JSON;
                var activityresults = responseReader.RetrieveData<Office365ActiveUsersServices>(reportingFilters);
                var results = activityresults.value.ToList();
                var rowprocessed = results.Count();

                // Switch to CSV Output
                reportingFilters.FormattedOutput = ReportUsageFormatEnum.Default;
                CSVConfig.RegisterClassMap<Office365ActiveUsersServicesMap>();
                var resultscsv = new CsvReader(responseReader.RetrieveDataAsStream(reportingFilters), CSVConfig);
                var resultsrec = resultscsv.GetRecords<Office365ActiveUsersServices>().ToList();
            }
            else if (reportingFilters.O365ReportType == ReportUsageTypeEnum.getOffice365ActiveUserCounts)
            {
                // Switch to JSON output
                reportingFilters.FormattedOutput = ReportUsageFormatEnum.JSON;
                var activityresults = responseReader.RetrieveData<Office365ActiveUsersUser>(reportingFilters);
                var results = activityresults.value.ToList();
                var rowprocessed = results.Count();

                // Switch to CSV Output
                reportingFilters.FormattedOutput = ReportUsageFormatEnum.Default;
                CSVConfig.RegisterClassMap<Office365ActiveUsersUserMap>();
                var resultscsv = new CsvReader(responseReader.RetrieveDataAsStream(reportingFilters), CSVConfig);
                var resultsrec = resultscsv.GetRecords<Office365ActiveUsersUser>().ToList();
            }
            else if (reportingFilters.O365ReportType == ReportUsageTypeEnum.getSharePointActivityFileCounts)
            {
                // Switch to JSON output
                reportingFilters.FormattedOutput = ReportUsageFormatEnum.JSON;
                var activityresults = responseReader.RetrieveData<SharePointActivityFileCounts>(reportingFilters);
                var results = activityresults.value.ToList();
                var rowprocessed = results.Count();

                // Switch to CSV Output
                reportingFilters.FormattedOutput = ReportUsageFormatEnum.Default;
                CSVConfig.RegisterClassMap<SharePointActivityFileCountsMap>();
                var resultscsv = new CsvReader(responseReader.RetrieveDataAsStream(reportingFilters), CSVConfig);
                var resultsrec = resultscsv.GetRecords<SharePointActivityFileCounts>().ToList();
            }
            else if (reportingFilters.O365ReportType == ReportUsageTypeEnum.getSharePointActivityPages)
            {
                // Switch to JSON output
                reportingFilters.FormattedOutput = ReportUsageFormatEnum.JSON;
                var activityresults = responseReader.RetrieveData<SharePointActivityPages>(reportingFilters);
                var results = activityresults.value.ToList();
                var rowprocessed = results.Count();

                // Switch to CSV Output
                reportingFilters.FormattedOutput = ReportUsageFormatEnum.Default;
                CSVConfig.RegisterClassMap<SharePointActivityPagesMap>();
                var resultscsv = new CsvReader(responseReader.RetrieveDataAsStream(reportingFilters), CSVConfig);
                var resultsrec = resultscsv.GetRecords<SharePointActivityPages>().ToList();
            }
            else if (reportingFilters.O365ReportType == ReportUsageTypeEnum.getSharePointActivityUserCounts)
            {
                // Switch to JSON output
                reportingFilters.FormattedOutput = ReportUsageFormatEnum.JSON;
                var activityresults = responseReader.RetrieveData<SharePointActivityUserCounts>(reportingFilters);
                var results = activityresults.value.ToList();
                var rowprocessed = results.Count();

                // Switch to CSV Output
                reportingFilters.FormattedOutput = ReportUsageFormatEnum.Default;
                CSVConfig.RegisterClassMap<SharePointActivityUserCountsMap>();
                var resultscsv = new CsvReader(responseReader.RetrieveDataAsStream(reportingFilters), CSVConfig);
                var resultsrec = resultscsv.GetRecords<SharePointActivityUserCounts>().ToList();
            }
            else if (reportingFilters.O365ReportType == ReportUsageTypeEnum.getSharePointActivityUserDetail)
            {
                // Switch to JSON output
                reportingFilters.FormattedOutput = ReportUsageFormatEnum.JSON;
                var activityresults = responseReader.RetrieveData<SharePointActivityUserDetail>(reportingFilters);
                var results = activityresults.value.ToList();
                var rowprocessed = results.Count();

                // Switch to CSV Output
                reportingFilters.FormattedOutput = ReportUsageFormatEnum.Default;
                CSVConfig.RegisterClassMap<SharePointActivityUserDetailMap>();
                var resultscsv = new CsvReader(responseReader.RetrieveDataAsStream(reportingFilters), CSVConfig);
                var resultsrec = resultscsv.GetRecords<SharePointActivityUserDetail>().ToList();
            }
            else if (reportingFilters.O365ReportType == ReportUsageTypeEnum.getSharePointSiteUsageDetail)
            {
                // Switch to JSON output
                reportingFilters.FormattedOutput = ReportUsageFormatEnum.JSON;
                var activityresults = responseReader.RetrieveData<SharePointSiteUsageDetail>(reportingFilters);
                var results = activityresults.value.ToList();
                var rowprocessed = results.Count();

                // Switch to CSV Output
                reportingFilters.FormattedOutput = ReportUsageFormatEnum.Default;
                CSVConfig.RegisterClassMap<SharePointSiteUsageDetailMap>();
                var resultscsv = new CsvReader(responseReader.RetrieveDataAsStream(reportingFilters), CSVConfig);
                var resultsrec = resultscsv.GetRecords<SharePointSiteUsageDetail>().ToList();
            }
            else if (reportingFilters.O365ReportType == ReportUsageTypeEnum.getSharePointSiteUsageFileCounts)
            {
                // Switch to JSON output
                reportingFilters.FormattedOutput = ReportUsageFormatEnum.JSON;
                var activityresults = responseReader.RetrieveData<SharePointSiteUsageFileCounts>(reportingFilters);
                var results = activityresults.value.ToList();
                var rowprocessed = results.Count();

                // Switch to CSV Output
                reportingFilters.FormattedOutput = ReportUsageFormatEnum.Default;
                CSVConfig.RegisterClassMap<SharePointSiteUsageFileCountsMap>();
                var resultscsv = new CsvReader(responseReader.RetrieveDataAsStream(reportingFilters), CSVConfig);
                var resultsrec = resultscsv.GetRecords<SharePointSiteUsageFileCounts>().ToList();
            }
            else if (reportingFilters.O365ReportType == ReportUsageTypeEnum.getSharePointSiteUsagePages)
            {
                // Switch to JSON output
                reportingFilters.FormattedOutput = ReportUsageFormatEnum.JSON;
                var activityresults = responseReader.RetrieveData<SharePointSiteUsagePages>(reportingFilters);
                var results = activityresults.value.ToList();
                var rowprocessed = results.Count();

                // Switch to CSV Output
                reportingFilters.FormattedOutput = ReportUsageFormatEnum.Default;
                CSVConfig.RegisterClassMap<SharePointSiteUsagePagesMap>();
                var resultscsv = new CsvReader(responseReader.RetrieveDataAsStream(reportingFilters), CSVConfig);
                var resultsrec = resultscsv.GetRecords<SharePointSiteUsagePages>().ToList();
            }
            else if (reportingFilters.O365ReportType == ReportUsageTypeEnum.getSharePointSiteUsageSiteCounts)
            {
                // Switch to JSON output
                reportingFilters.FormattedOutput = ReportUsageFormatEnum.JSON;
                var activityresults = responseReader.RetrieveData<SharePointSiteUsageSiteCounts>(reportingFilters);
                var results = activityresults.value.ToList();
                var rowprocessed = results.Count();

                // Switch to CSV Output
                reportingFilters.FormattedOutput = ReportUsageFormatEnum.Default;
                CSVConfig.RegisterClassMap<SharePointSiteUsageSiteCountsMap>();
                var resultscsv = new CsvReader(responseReader.RetrieveDataAsStream(reportingFilters), CSVConfig);
                var resultsrec = resultscsv.GetRecords<SharePointSiteUsageSiteCounts>().ToList();
            }
            else if (reportingFilters.O365ReportType == ReportUsageTypeEnum.getSharePointSiteUsageStorage)
            {
                // Switch to JSON output
                reportingFilters.FormattedOutput = ReportUsageFormatEnum.JSON;
                var activityresults = responseReader.RetrieveData<SharePointSiteUsageStorage>(reportingFilters);
                var results = activityresults.value.ToList();
                var rowprocessed = results.Count();

                // Switch to CSV Output
                reportingFilters.FormattedOutput = ReportUsageFormatEnum.Default;
                CSVConfig.RegisterClassMap<SharePointSiteUsageStorageMap>();
                var resultscsv = new CsvReader(responseReader.RetrieveDataAsStream(reportingFilters), CSVConfig);
                var resultsrec = resultscsv.GetRecords<SharePointSiteUsageStorage>().ToList();
            }
            else
            {
                var response = responseReader.RetrieveData(reportingFilters);
                Logger.LogInformation("WebResponse:{0}", response);
            }
        }


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