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

        [Parameter(Mandatory = true, ValueFromPipeline = true, ValueFromPipelineByPropertyName = true, Position = 4)]
        public ReportUsageTypeEnum ReportType { get; set; }

        [Parameter(Mandatory = false, ValueFromPipeline = true, ValueFromPipelineByPropertyName = true, Position = 5)]
        public ReportUsagePeriodEnum Period { get; set; }

        [Parameter(Mandatory = false, ValueFromPipeline = true, ValueFromPipelineByPropertyName = true, Position = 6)]
        public Nullable<DateTime> Date { get; set; }


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

            var filter = new QueryFilter()
            {
                O365Period = Period,
                O365ReportType = ReportType,
                Date = Date
            };

            // CSV config to process in memory
            var csvconfig = new Configuration()
            {
                Delimiter = ",",
                HasHeaderRecord = true
            };


            var ilogger = new DefaultUsageLogger(LogVerbose, LogWarning, LogError);
            ilogger.LogInformation("Report => Usage Type {0} Period {1}", ReportType, Period);


            using (var reporter = new ExampleReportVisitor(csvconfig, filter, ilogger))
            {
                ReportingStream stream = new ReportingStream(filter, config, ilogger);
                stream.RetrieveData(reporter);
            }
        }
    }
}
