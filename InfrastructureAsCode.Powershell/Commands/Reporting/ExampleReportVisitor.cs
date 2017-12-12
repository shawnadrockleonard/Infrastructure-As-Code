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
                CSVConfig.RegisterClassMap<OneDriveActivityFileCountsMap>();
                var activityfilescsv = new CsvReader(responseReader.RetrieveDataAsStream(reportingFilters), CSVConfig);
                var activityfilesresults = activityfilescsv.GetRecords<OneDriveActivityFileCounts>().ToList();
            }
            else if (reportingFilters.O365ReportType == ReportUsageTypeEnum.getOneDriveActivityUserCounts)
            {
                CSVConfig.RegisterClassMap<OneDriveActivityUserCountsMap>();
                var activityuserscsv = new CsvReader(responseReader.RetrieveDataAsStream(reportingFilters), CSVConfig);
                var activityusersresults = activityuserscsv.GetRecords<OneDriveActivityUserCounts>().ToList();
            }
            else if (reportingFilters.O365ReportType == ReportUsageTypeEnum.getOneDriveActivityUserDetail)
            {
                CSVConfig.RegisterClassMap<OneDriveActivityDetailMap>();
                var activitydetailcsv = new CsvReader(responseReader.RetrieveDataAsStream(reportingFilters), CSVConfig);
                var activitydetailresults = activitydetailcsv.GetRecords<OneDriveActivityDetail>().ToList();
            }
            else if (reportingFilters.O365ReportType == ReportUsageTypeEnum.getOneDriveUsageStorage)
            {
                CSVConfig.RegisterClassMap<OneDriveUsageStorageMap>();
                var resultscsv = new CsvReader(responseReader.RetrieveDataAsStream(reportingFilters), CSVConfig);
                var results = resultscsv.GetRecords<OneDriveUsageStorage>().ToList();
            }
            else if (reportingFilters.O365ReportType == ReportUsageTypeEnum.getOneDriveUsageFileCounts)
            {
                CSVConfig.RegisterClassMap<OneDriveUsageFileCountsMap>();
                var resultscsv = new CsvReader(responseReader.RetrieveDataAsStream(reportingFilters), CSVConfig);
                var results = resultscsv.GetRecords<OneDriveUsageFileCounts>().ToList();
            }
            else if (reportingFilters.O365ReportType == ReportUsageTypeEnum.getOneDriveUsageAccountCounts)
            {
                CSVConfig.RegisterClassMap<OneDriveUsageAccountCountsMap>();
                var resultscsv = new CsvReader(responseReader.RetrieveDataAsStream(reportingFilters), CSVConfig);
                var results = resultscsv.GetRecords<OneDriveUsageAccountCounts>().ToList();
            }
            else if (reportingFilters.O365ReportType == ReportUsageTypeEnum.getOneDriveUsageAccountDetail)
            {
                CSVConfig.RegisterClassMap<OneDriveUsageAccountDetailMap>();
                var resultscsv = new CsvReader(responseReader.RetrieveDataAsStream(reportingFilters), CSVConfig);
                var results = resultscsv.GetRecords<OneDriveUsageAccountDetail>().ToList();
            }
            else if (reportingFilters.O365ReportType == ReportUsageTypeEnum.getOffice365ActiveUserDetail)
            {
                CSVConfig.RegisterClassMap<Office365ActiveUsersDetailsMap>();
                var resultscsv = new CsvReader(responseReader.RetrieveDataAsStream(reportingFilters), CSVConfig);
                var results = resultscsv.GetRecords<Office365ActiveUsersDetails>().ToList();
            }
            else if (reportingFilters.O365ReportType == ReportUsageTypeEnum.getOffice365ServicesUserCounts)
            {
                CSVConfig.RegisterClassMap<Office365ActiveUsersServicesMap>();
                var resultscsv = new CsvReader(responseReader.RetrieveDataAsStream(reportingFilters), CSVConfig);
                var results = resultscsv.GetRecords<Office365ActiveUsersServices>().ToList();
            }
            else if (reportingFilters.O365ReportType == ReportUsageTypeEnum.getOffice365ActiveUserCounts)
            {
                CSVConfig.RegisterClassMap<Office365ActiveUsersUserMap>();
                var resultscsv = new CsvReader(responseReader.RetrieveDataAsStream(reportingFilters), CSVConfig);
                var results = resultscsv.GetRecords<Office365ActiveUsersUser>().ToList();
            }
            else if (reportingFilters.O365ReportType == ReportUsageTypeEnum.getSharePointActivityFileCounts)
            {
                CSVConfig.RegisterClassMap<SharePointActivityFileCountsMap>();
                var resultscsv = new CsvReader(responseReader.RetrieveDataAsStream(reportingFilters), CSVConfig);
                var results = resultscsv.GetRecords<SharePointActivityFileCounts>().ToList();
            }
            else if (reportingFilters.O365ReportType == ReportUsageTypeEnum.getSharePointActivityPages)
            {
                CSVConfig.RegisterClassMap<SharePointActivityPagesMap>();
                var resultscsv = new CsvReader(responseReader.RetrieveDataAsStream(reportingFilters), CSVConfig);
                var results = resultscsv.GetRecords<SharePointActivityPages>().ToList();
            }
            else if (reportingFilters.O365ReportType == ReportUsageTypeEnum.getSharePointActivityUserCounts)
            {
                CSVConfig.RegisterClassMap<SharePointActivityUserCountsMap>();
                var resultscsv = new CsvReader(responseReader.RetrieveDataAsStream(reportingFilters), CSVConfig);
                var results = resultscsv.GetRecords<SharePointActivityUserCounts>().ToList();
            }
            else if (reportingFilters.O365ReportType == ReportUsageTypeEnum.getSharePointActivityUserDetail)
            {
                CSVConfig.RegisterClassMap<SharePointActivityUserDetailMap>();
                var resultscsv = new CsvReader(responseReader.RetrieveDataAsStream(reportingFilters), CSVConfig);
                var results = resultscsv.GetRecords<SharePointActivityUserDetail>().ToList();
            }
            else if (reportingFilters.O365ReportType == ReportUsageTypeEnum.getSharePointSiteUsageDetail)
            {
                CSVConfig.RegisterClassMap<SharePointSiteUsageDetailMap>();
                var resultscsv = new CsvReader(responseReader.RetrieveDataAsStream(reportingFilters), CSVConfig);
                var results = resultscsv.GetRecords<SharePointSiteUsageDetail>().ToList();
            }
            else if (reportingFilters.O365ReportType == ReportUsageTypeEnum.getSharePointSiteUsageFileCounts)
            {
                CSVConfig.RegisterClassMap<SharePointSiteUsageFileCountsMap>();
                var resultscsv = new CsvReader(responseReader.RetrieveDataAsStream(reportingFilters), CSVConfig);
                var results = resultscsv.GetRecords<SharePointSiteUsageFileCounts>().ToList();
            }
            else if (reportingFilters.O365ReportType == ReportUsageTypeEnum.getSharePointSiteUsagePages)
            {
                CSVConfig.RegisterClassMap<SharePointSiteUsagePagesMap>();
                var resultscsv = new CsvReader(responseReader.RetrieveDataAsStream(reportingFilters), CSVConfig);
                var results = resultscsv.GetRecords<SharePointSiteUsagePages>().ToList();
            }
            else if (reportingFilters.O365ReportType == ReportUsageTypeEnum.getSharePointSiteUsageSiteCounts)
            {
                CSVConfig.RegisterClassMap<SharePointSiteUsageSiteCountsMap>();
                var resultscsv = new CsvReader(responseReader.RetrieveDataAsStream(reportingFilters), CSVConfig);
                var results = resultscsv.GetRecords<SharePointSiteUsageSiteCounts>().ToList();
            }
            else if (reportingFilters.O365ReportType == ReportUsageTypeEnum.getSharePointSiteUsageStorage)
            {
                CSVConfig.RegisterClassMap<SharePointSiteUsageStorageMap>();
                var resultscsv = new CsvReader(responseReader.RetrieveDataAsStream(reportingFilters), CSVConfig);
                var results = resultscsv.GetRecords<SharePointSiteUsageStorage>().ToList();
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