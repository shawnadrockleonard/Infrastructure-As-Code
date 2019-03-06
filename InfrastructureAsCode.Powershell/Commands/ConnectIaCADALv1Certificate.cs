using InfrastructureAsCode.Core;
using InfrastructureAsCode.Core.oAuth;
using InfrastructureAsCode.Powershell.Commands.Base;
using OfficeDevPnP.Core;
using System;
using System.Linq;
using System.Management.Automation;
using System.Security.Cryptography.X509Certificates;

namespace InfrastructureAsCode.Powershell.Commands
{
    [Cmdlet("Connect", "IaCADALv1Certificate")]
    public class ConnectIaCADALv1Certificate : ExtendedPSCmdlet
    {
        [Parameter(Mandatory = true, HelpMessage = "The client id of the app which gives you access to the Microsoft Graph API.", ParameterSetName = "AAD")]
        public string AppId { get; set; }

        [Parameter(Mandatory = true, HelpMessage = "The app key of the app which gives you access to the Microsoft Graph API.", ParameterSetName = "AAD")]
        public string Thumbprint { get; set; }

        [Parameter(Mandatory = true, HelpMessage = "The AAD where the O365 app is registred. Eg.: contoso.com, or contoso.onmicrosoft.com.", ParameterSetName = "AAD")]
        public string TenantDomain { get; set; }

        [Parameter(Mandatory = false, HelpMessage = "The AAD where the O365 app is registered. Eg.: {guid}", ParameterSetName = "AAD")]
        public string TenantId { get; set; }

        [Parameter(Mandatory = false, HelpMessage = "The URI of the resource to query", ParameterSetName = "AAD")]
        public string ResourceUri { get; set; }

        [Parameter(Mandatory = false, HelpMessage = "The AAD authority to which you are connecting", ParameterSetName = "AAD")]
        public AzureEnvironment Environment { get; set; }


        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();

            var config = new AzureADConfig()
            {
                ClientId = this.AppId,
                CertificateThumbprint = this.Thumbprint,
                Certificate = ReadCertificateFromStore(this.Thumbprint),
                RedirectUri = ResourceUri ?? AzureADConstants.O365DefaultId,
                AuthenticationEndpoint = Environment,
                TenantDomain = this.TenantDomain,
                TenantId = TenantId ?? ""
            };

            var ilogger = new DefaultUsageLogger(LogVerbose, LogWarning, LogError);


            var authenticationEndpoint = Environment.GetAzureADLoginEndPoint();
            var endpoint = string.Format(AzureADConstants.AuthorityTenantFormat, authenticationEndpoint, this.TenantDomain);

            AzureADALv1Connection.CurrentConnection = new AzureADALv1Connection(endpoint, config, ilogger);


            // Write Tokens to Console
            WriteObject(AzureADALv1Connection.CurrentConnection.AuthenticationResult);
        }


        private static X509Certificate2 ReadCertificateFromStore(string thumbprint)
        {
            X509Certificate2 cert = null;
            var store = new X509Store(StoreName.My, StoreLocation.CurrentUser);
            store.Open(OpenFlags.ReadOnly);
            X509Certificate2Collection certCollection = store.Certificates;
            X509Certificate2Collection currentCerts = certCollection.Find(X509FindType.FindByTimeValid, DateTime.Now, false);
            X509Certificate2Collection signingCert = currentCerts.Find(X509FindType.FindByThumbprint, thumbprint, false);
            cert = signingCert.OfType<X509Certificate2>().OrderByDescending(c => c.NotBefore).FirstOrDefault();
            store.Close();
            return cert;
        }
    }
}

