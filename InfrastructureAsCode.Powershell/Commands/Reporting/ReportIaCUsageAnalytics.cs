using InfrastructureAsCode.Core;
using InfrastructureAsCode.Core.oAuth;
using InfrastructureAsCode.Core.Reports;
using InfrastructureAsCode.Core.Reports.o365Graph;
using InfrastructureAsCode.Core.Reports.o365Graph.TenantReport.Mappings;
using InfrastructureAsCode.Powershell.CmdLets;
using System;
using System.Management.Automation;

namespace InfrastructureAsCode.Powershell.Commands.Reporting
{
    [Cmdlet(VerbsExtended.Report, "IaCUsageAnalytics", SupportsShouldProcess = false)]
    [CmdletHelp("Connects to a Azure AD to claim a token and process a usage report",
        DetailedDescription = "This is a sample for querying the preview MS Graph APIs.",
        Category = "Preview Reporting Cmdlets")]
    public class ReportIaCUsageAnalytics : IaCAdminCmdlet
    {
        [Parameter(Mandatory = true, ValueFromPipeline = true, ValueFromPipelineByPropertyName = true, Position = 4)]
        public ReportUsageTypeEnum ReportType { get; set; }

        [Parameter(Mandatory = true, ValueFromPipeline = true, ValueFromPipelineByPropertyName = true, Position = 5)]
        public ReportUsagePeriodEnum Period { get; set; }

        [Parameter(Mandatory = true, ValueFromPipeline = true, ValueFromPipelineByPropertyName = true, Position = 6)]
        public Nullable<DateTime> Date { get; set; }


        [Parameter(Mandatory = false, ValueFromPipeline = false)]
        public SwitchParameter BetaEndPoint { get; set; }

        [Parameter(Mandatory = false, ValueFromPipeline = false)]
        public string DataDirectory { get; set; }


        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();

            var defaultRows = 500;
            var _dateLog = DateTime.UtcNow;
            var _fileName = ReportType.ToString("f");
            var _logger = new DefaultUsageLogger(LogVerbose, LogWarning, LogError);
            _logger.LogInformation("Report => Usage Type {0} Period {1}", ReportType, Period);


            if (!string.IsNullOrEmpty(DataDirectory))
            {
                var logFileModel = new ReportDirectoryHandler(DataDirectory, _fileName, _dateLog, _logger);

                // Deleting log file;
                logFileModel.ResetCSVFile();

                var _reportingProcessor = new ReportingProcessor(this.AzureADConfig, _logger);
                var response = _reportingProcessor.ProcessReport(ReportType, Period, Date, defaultRows, BetaEndPoint);
                logFileModel.WriteToCSVFile(response);
            }
            else
            {

                var _reportingProcessor = new ReportingProcessor(this.AzureADConfig, _logger);



                var office365ActiveUsersUserDetail = _reportingProcessor.ProcessReport<Office365ActiveUsersUserDetail>(Period, Date, defaultRows, BetaEndPoint);
                var office365ActiveUsersServicesUserCounts = _reportingProcessor.ProcessReport<Office365ActiveUsersServicesUserCounts>(Period, Date, defaultRows, BetaEndPoint);
                var office365ActiveUsersUserCounts = _reportingProcessor.ProcessReport<Office365ActiveUsersUserCounts>(Period, Date, defaultRows, BetaEndPoint);

                var varOffice365GroupsActivityCounts = _reportingProcessor.ProcessReport<Office365GroupsActivityCounts>(Period, Date, defaultRows, BetaEndPoint);
                var varOffice365GroupsActivityDetail = _reportingProcessor.ProcessReport<Office365GroupsActivityDetail>(Period, Date, defaultRows, BetaEndPoint);

                var varSkypeForBusinessPeerToPeerActivityCounts = _reportingProcessor.ProcessReport<SkypeForBusinessPeerToPeerActivityCounts>(Period, Date, defaultRows, BetaEndPoint);
                var varSkypeForBusinessPeerToPeerActivityMinuteCounts = _reportingProcessor.ProcessReport<SkypeForBusinessPeerToPeerActivityMinuteCounts>(Period, Date, defaultRows, BetaEndPoint);
                var varSkypeForBusinessPeerToPeerActivityUserCounts = _reportingProcessor.ProcessReport<SkypeForBusinessPeerToPeerActivityUserCounts>(Period, Date, defaultRows, BetaEndPoint);
                var varSkypeForBusinessParticipantActivityCounts = _reportingProcessor.ProcessReport<SkypeForBusinessParticipantActivityCounts>(Period, Date, defaultRows, BetaEndPoint);
                var varSkypeForBusinessParticipantActivityMinuteCounts = _reportingProcessor.ProcessReport<SkypeForBusinessParticipantActivityMinuteCounts>(Period, Date, defaultRows, BetaEndPoint);
                var varSkypeForBusinessParticipantActivityUserCounts = _reportingProcessor.ProcessReport<SkypeForBusinessParticipantActivityUserCounts>(Period, Date, defaultRows, BetaEndPoint);
                var varSkypeForBusinessOrganizerActivityCounts = _reportingProcessor.ProcessReport<SkypeForBusinessOrganizerActivityCounts>(Period, Date, defaultRows, BetaEndPoint);
                var varSkypeForBusinessOrganizerActivityMinuteCounts = _reportingProcessor.ProcessReport<SkypeForBusinessOrganizerActivityMinuteCounts>(Period, Date, defaultRows, BetaEndPoint);
                var varSkypeForBusinessOrganizerActivityUserCounts = _reportingProcessor.ProcessReport<SkypeForBusinessOrganizerActivityUserCounts>(Period, Date, defaultRows, BetaEndPoint);
                var varSkypeForBusinessDeviceUsageDistributionUserCounts = _reportingProcessor.ProcessReport<SkypeForBusinessDeviceUsageDistributionUserCounts>(Period, Date, defaultRows, BetaEndPoint);
                var varSkypeForBusinessDeviceUsageUserDetail = _reportingProcessor.ProcessReport<SkypeForBusinessDeviceUsageUserDetail>(Period, Date, defaultRows, BetaEndPoint);
                var varSkypeForBusinessDeviceUsageUserCounts = _reportingProcessor.ProcessReport<SkypeForBusinessDeviceUsageUserCounts>(Period, Date, defaultRows, BetaEndPoint);
                var skypeForBusinessActivityUserDetail = _reportingProcessor.ProcessReport<SkypeForBusinessActivityUserDetail>(Period, Date, defaultRows, BetaEndPoint);
                var skypeForBusinessActivityCounts = _reportingProcessor.ProcessReport<SkypeForBusinessActivityActivityCounts>(Period, Date, defaultRows, BetaEndPoint);
                var skypeForBusinessActivityUserCounts = _reportingProcessor.ProcessReport<SkypeForBusinessActivityUserCounts>(Period, Date, defaultRows, BetaEndPoint);


                var onedriveactivityuserdetail = _reportingProcessor.ProcessReport<OneDriveActivityUserDetail>(Period, Date, defaultRows, BetaEndPoint);
                var onedriveactivityusercounts = _reportingProcessor.ProcessReport<OneDriveActivityUserCounts>(Period, Date, defaultRows, BetaEndPoint);
                var onedriveactivityfilecounts = _reportingProcessor.ProcessReport<OneDriveActivityFileCounts>(Period, Date, defaultRows, BetaEndPoint);

                var onedriveausageaccountdetail = _reportingProcessor.ProcessReport<OneDriveUsageAccountDetail>(Period, Date, defaultRows, BetaEndPoint);
                var onedriveusageaccountcounts = _reportingProcessor.ProcessReport<OneDriveUsageAccountCounts>(Period, Date, defaultRows, BetaEndPoint);
                var onedriveusagefilecounts = _reportingProcessor.ProcessReport<OneDriveUsageFileCounts>(Period, Date, defaultRows, BetaEndPoint);
                var onedriveusagestorage = _reportingProcessor.ProcessReport<OneDriveUsageStorage>(Period, Date, defaultRows, BetaEndPoint);


                var sharePointActivityUserDetail = _reportingProcessor.ProcessReport<SharePointActivityUserDetail>(Period, Date, defaultRows, BetaEndPoint);
                var sharePointActivityFileCounts = _reportingProcessor.ProcessReport<SharePointActivityFileCounts>(Period, Date, defaultRows, BetaEndPoint);
                var sharePointActivityUserCounts = _reportingProcessor.ProcessReport<SharePointActivityUserCounts>(Period, Date, defaultRows, BetaEndPoint);
                var sharePointActivityPages = _reportingProcessor.ProcessReport<SharePointActivityPages>(Period, Date, defaultRows, BetaEndPoint);

                var sharePointSiteUsageSiteDetail = _reportingProcessor.ProcessReport<SharePointSiteUsageSiteDetail>(Period, Date, defaultRows, BetaEndPoint);
                var sharePointSiteUsageFileCounts = _reportingProcessor.ProcessReport<SharePointSiteUsageFileCounts>(Period, Date, defaultRows, BetaEndPoint);
                var sharePointSiteUsageSiteCounts = _reportingProcessor.ProcessReport<SharePointSiteUsageSiteCounts>(Period, Date, defaultRows, BetaEndPoint);
                var sharePointSiteUsageStorage = _reportingProcessor.ProcessReport<SharePointSiteUsageStorage>(Period, Date, defaultRows, BetaEndPoint);
                var sharePointSiteUsagePages = _reportingProcessor.ProcessReport<SharePointSiteUsagePages>(Period, Date, defaultRows, BetaEndPoint);

            }
        }
    }
}
