using OfficeDevPnP.Core;
using System.Security.Cryptography.X509Certificates;

namespace InfrastructureAsCode.Core.oAuth
{
    /// <summary>
    /// Represents the config for claiming and refreshing tokens
    /// </summary>
    public class AzureADConfig : IAzureADConfig
    {
        public string CallbackPath { get; set; }

        public string ClientId { get; set; }

        public string ClientSecret { get; set; }

        public string CertificateThumbprint { get; set; }

        public X509Certificate2 Certificate { get; set; }

        public string RedirectUri { get; set; }

        public AzureEnvironment AuthenticationEndpoint { get; set; }

        public string TenantDomain { get; set; }

        public string TenantId { get; set; }

        public string SecurityGroupId { get; set; }

        public bool IsCertificateAuth()
        {
            return !string.IsNullOrEmpty(CertificateThumbprint) || Certificate != null;
        }
    }
}
