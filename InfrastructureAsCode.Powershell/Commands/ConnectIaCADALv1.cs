using InfrastructureAsCode.Core;
using InfrastructureAsCode.Core.oAuth;
using InfrastructureAsCode.Powershell.CmdLets;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System;
using System.Management.Automation;
using System.Threading.Tasks;


namespace InfrastructureAsCode.Powershell.Commands
{
    [Cmdlet("Connect", "IaCADALv1")]
    public class ConnectIaCADALv1 : ExtendedPSCmdlet
    {
        [Parameter(Mandatory = true, HelpMessage = "The client id of the app which gives you access to the Microsoft Graph API.", ParameterSetName = "AAD")]
        public string AppId { get; set; }

        [Parameter(Mandatory = true, HelpMessage = "The app key of the app which gives you access to the Microsoft Graph API.", ParameterSetName = "AAD")]
        public string AppSecret { get; set; }

        [Parameter(Mandatory = true, HelpMessage = "The AAD where the O365 app is registred. Eg.: contoso.com, or contoso.onmicrosoft.com.", ParameterSetName = "AAD")]
        public string AADDomain { get; set; }

        [Parameter(Mandatory = true, HelpMessage = "The URI of the resource to query", ParameterSetName = "AAD")]
        public string ResourceUri { get; set; }


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


            AzureADALv1Connection.CurrentConnection = new AzureADALv1Connection(config, ilogger);


            // Write Tokens to Console
            WriteObject(AzureADALv1Connection.CurrentConnection.AuthenticationResult);
        }

    }
}

