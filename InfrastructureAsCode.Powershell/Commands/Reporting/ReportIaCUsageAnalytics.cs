using InfrastructureAsCode.Core.Reports;
using InfrastructureAsCode.Core.Reports.o365Graph;
using InfrastructureAsCode.Core.Reports.o365Graph.AzureAD;
using InfrastructureAsCode.Powershell.CmdLets;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;
using CsvHelper;
using CsvHelper.Configuration;
using InfrastructureAsCode.Core.Reports.o365Graph.TenantReport;
using InfrastructureAsCode.Core.Reports.o365Graph.TenantReport.Mappings;

namespace InfrastructureAsCode.Powershell.Commands.Reporting
{
    [Cmdlet(VerbsExtended.Report, "IaCUsageAnalytics", SupportsShouldProcess = false)]
    [CmdletHelp("Connects to a Azure AD to claim a token and process a usage report",
        DetailedDescription = "This is a sample for querying the preview MS Graph APIs.",
        Category = "Preview Reporting Cmdlets")]
    public class ReportIaCUsageAnalytics : ExtendedPSCmdlet
    {
        private const string RedirectUri = "urn:ietf:wg:oauth:2.0:oob";

        [Parameter(Mandatory = true, HelpMessage = "The client id of the app which gives you access to the Microsoft Graph API.", ParameterSetName = "AAD")]
        public string AppId { get; set; }

        [Parameter(Mandatory = true, HelpMessage = "The app key of the app which gives you access to the Microsoft Graph API.", ParameterSetName = "AAD")]
        public string AppSecret { get; set; }

        [Parameter(Mandatory = true, HelpMessage = "The AAD where the O365 app is registred. Eg.: contoso.com, or contoso.onmicrosoft.com.", ParameterSetName = "AAD")]
        public string AADDomain { get; set; }

        [Parameter(Mandatory = true, HelpMessage = "The URI of the resource to query", ParameterSetName = "AAD")]
        public string ResourceUri { get; set; }

        [Parameter(Mandatory = false, ValueFromPipeline = true, ValueFromPipelineByPropertyName = true, Position = 4)]
        public ReportUsageTypeEnum ReportType { get; set; }

        [Parameter(Mandatory = false, ValueFromPipeline = true, ValueFromPipelineByPropertyName = true, Position = 5)]
        public ReportUsagePeriodEnum Period { get; set; }

        [Parameter(Mandatory = false, ValueFromPipeline = true, ValueFromPipelineByPropertyName = true, Position = 6)]
        public Nullable<DateTime> Date { get; set; }

        [Parameter(Mandatory = false)]
        public SwitchParameter BetaEndPoint { get; set; }



        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();

            var config = new AzureADConfig()
            {
                ClientId = this.AppId,
                ClientSecret = this.AppSecret,
                RedirectUri = AzureADConstants.GraphResourceId,
                TenantDomain = this.AADDomain,
                TenantId = ""
            };


            var ilogger = new DefaultUsageLogger(LogVerbose, LogWarning, LogError);
            ilogger.LogInformation("Report => Usage Type {0} Period {1}", ReportType, Period);

            var reporter = new ReportingProcessor(config, ilogger);

            var overrideDate = new DateTime(2018, 5, 30);


            var varOffice365GroupsActivityCounts = reporter.ProcessReport<Office365GroupsActivityCounts>(Period, Date, 500, BetaEndPoint);
            var varOffice365GroupsActivityCountscsv = reporter.ProcessReport<Office365GroupsActivityCounts>(Period, Date, 500, false);
            var varOffice365GroupsActivityDetail = reporter.ProcessReport<Office365GroupsActivityDetail>(Period, Date, 500, BetaEndPoint);


            var varSkypeForBusinessPeerToPeerActivityCounts = reporter.ProcessReport<SkypeForBusinessPeerToPeerActivityCounts>(Period, Date, 500, BetaEndPoint);
            var varSkypeForBusinessPeerToPeerActivityMinuteCounts = reporter.ProcessReport<SkypeForBusinessPeerToPeerActivityMinuteCounts>(Period, Date, 500, BetaEndPoint);
            var varSkypeForBusinessPeerToPeerActivityUserCounts = reporter.ProcessReport<SkypeForBusinessPeerToPeerActivityUserCounts>(Period, Date, 500, BetaEndPoint);

            var varSkypeForBusinessPeerToPeerActivityCountscsv = reporter.ProcessReport<SkypeForBusinessPeerToPeerActivityCounts>(Period, Date, 500, false);
            var varSkypeForBusinessPeerToPeerActivityMinuteCountscsv = reporter.ProcessReport<SkypeForBusinessPeerToPeerActivityMinuteCounts>(Period, Date, 500, false);
            var varSkypeForBusinessPeerToPeerActivityUserCountscsv = reporter.ProcessReport<SkypeForBusinessPeerToPeerActivityUserCounts>(Period, Date, 500, false);



            var varSkypeForBusinessParticipantActivityCounts = reporter.ProcessReport<SkypeForBusinessParticipantActivityCounts>(Period, Date, 500, BetaEndPoint);
            var varSkypeForBusinessParticipantActivityMinuteCounts = reporter.ProcessReport<SkypeForBusinessParticipantActivityMinuteCounts>(Period, Date, 500, BetaEndPoint);
            var varSkypeForBusinessParticipantActivityUserCounts = reporter.ProcessReport<SkypeForBusinessParticipantActivityUserCounts>(Period, Date, 500, BetaEndPoint);

            var varSkypeForBusinessParticipantActivityCountscsv = reporter.ProcessReport<SkypeForBusinessParticipantActivityCounts>(Period, Date, 500, false);
            var varSkypeForBusinessParticipantActivityMinuteCountscsv = reporter.ProcessReport<SkypeForBusinessParticipantActivityMinuteCounts>(Period, Date, 500, false);
            var varSkypeForBusinessParticipantActivityUserCountscsv = reporter.ProcessReport<SkypeForBusinessParticipantActivityUserCounts>(Period, Date, 500, false);


            var varSkypeForBusinessOrganizerActivityCounts = reporter.ProcessReport<SkypeForBusinessOrganizerActivityCounts>(Period, Date, 500, BetaEndPoint);
            var varSkypeForBusinessOrganizerActivityMinuteCounts = reporter.ProcessReport<SkypeForBusinessOrganizerActivityMinuteCounts>(Period, Date, 500, BetaEndPoint);
            var varSkypeForBusinessOrganizerActivityUserCounts = reporter.ProcessReport<SkypeForBusinessOrganizerActivityUserCounts>(Period, Date, 500, BetaEndPoint);

            var varSkypeForBusinessOrganizerActivityCountscsv = reporter.ProcessReport<SkypeForBusinessOrganizerActivityCounts>(Period, Date, 500, false);
            var varSkypeForBusinessOrganizerActivityMinuteCountscsv = reporter.ProcessReport<SkypeForBusinessOrganizerActivityMinuteCounts>(Period, Date, 500, false);
            var varSkypeForBusinessOrganizerActivityUserCountscsv = reporter.ProcessReport<SkypeForBusinessOrganizerActivityUserCounts>(Period, Date, 500, false);

            var varSkypeForBusinessDeviceUsageDistributionUserCounts = reporter.ProcessReport<SkypeForBusinessDeviceUsageDistributionUserCounts>(Period, Date, 500, BetaEndPoint);
            var varSkypeForBusinessDeviceUsageDistributionUserCountscsv = reporter.ProcessReport<SkypeForBusinessDeviceUsageDistributionUserCounts>(Period, Date, 500, false);
            var varSkypeForBusinessDeviceUsageUserDetail = reporter.ProcessReport<SkypeForBusinessDeviceUsageUserDetail>(Period, overrideDate, 500, BetaEndPoint);
            var varSkypeForBusinessDeviceUsageUserDetailcsv = reporter.ProcessReport<SkypeForBusinessDeviceUsageUserDetail>(Period, overrideDate, 500, false);
            var varSkypeForBusinessDeviceUsageUserCounts = reporter.ProcessReport<SkypeForBusinessDeviceUsageUserCounts>(Period, overrideDate, 500, BetaEndPoint);
            var varSkypeForBusinessDeviceUsageUserCountscsv = reporter.ProcessReport<SkypeForBusinessDeviceUsageUserCounts>(Period, overrideDate, 500, false);


            var skypeForBusinessActivityUserDetail = reporter.ProcessReport<SkypeForBusinessActivityUserDetail>(Period, overrideDate, 500, BetaEndPoint);
            var skypeForBusinessActivityCounts = reporter.ProcessReport<SkypeForBusinessActivityActivityCounts>(Period, Date, 500, BetaEndPoint);
            var skypeForBusinessActivityUserCounts = reporter.ProcessReport<SkypeForBusinessActivityUserCounts>(Period, overrideDate, 500, BetaEndPoint);


            var onedriveactivityuserdetail = reporter.ProcessReport<OneDriveActivityUserDetail>(Period, overrideDate, 500, BetaEndPoint);
            var onedriveactivityusercounts = reporter.ProcessReport<OneDriveActivityUserCounts>(Period, Date, 500, BetaEndPoint);
            var onedriveactivityfilecounts = reporter.ProcessReport<OneDriveActivityFileCounts>(Period, Date, 500, BetaEndPoint);

            var onedriveausageaccountdetail = reporter.ProcessReport<OneDriveUsageAccountDetail>(Period, Date, 500, BetaEndPoint);
            var onedriveusageaccountcounts = reporter.ProcessReport<OneDriveUsageAccountCounts>(Period, Date, 500, BetaEndPoint);
            var onedriveusagefilecounts = reporter.ProcessReport<OneDriveUsageFileCounts>(Period, Date, 500, BetaEndPoint);
            var onedriveusagestorage = reporter.ProcessReport<OneDriveUsageStorage>(Period, Date, 500, BetaEndPoint);



            var office365ActiveUsersUserDetail = reporter.ProcessReport<Office365ActiveUsersUserDetail>(Period, Date, 500, BetaEndPoint);
            var office365ActiveUsersServicesUserCounts = reporter.ProcessReport<Office365ActiveUsersServicesUserCounts>(Period, Date, 500, BetaEndPoint);
            var office365ActiveUsersUserCounts = reporter.ProcessReport<Office365ActiveUsersUserCounts>(Period, Date, 500, BetaEndPoint);



            var sharePointActivityUserDetail = reporter.ProcessReport<SharePointActivityUserDetail>(Period, Date, 500, BetaEndPoint);
            var sharePointActivityFileCounts = reporter.ProcessReport<SharePointActivityFileCounts>(Period, Date, 500, BetaEndPoint);
            var sharePointActivityUserCounts = reporter.ProcessReport<SharePointActivityUserCounts>(Period, Date, 500, BetaEndPoint);
            var sharePointActivityPages = reporter.ProcessReport<SharePointActivityPages>(Period, Date, 500, BetaEndPoint);

            var sharePointSiteUsageSiteDetail = reporter.ProcessReport<SharePointSiteUsageSiteDetail>(Period, Date, 500, BetaEndPoint);
            var sharePointSiteUsageFileCounts = reporter.ProcessReport<SharePointSiteUsageFileCounts>(Period, Date, 500, BetaEndPoint);
            var sharePointSiteUsageSiteCounts = reporter.ProcessReport<SharePointSiteUsageSiteCounts>(Period, Date, 500, BetaEndPoint);
            var sharePointSiteUsageStorage = reporter.ProcessReport<SharePointSiteUsageStorage>(Period, Date, 500, BetaEndPoint);
            var sharePointSiteUsagePages = reporter.ProcessReport<SharePointSiteUsagePages>(Period, Date, 500, BetaEndPoint);





        }
    }
}
