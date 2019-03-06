using OfficeDevPnP.Core;
using System.Security.Cryptography.X509Certificates;

namespace InfrastructureAsCode.Core.oAuth
{
    /// <summary>
    /// Azure AD Config settings for claiming tokens
    /// </summary>
    public interface IAzureADConfig
    {

        string CallbackPath { get; set; }

        string ClientId { get; set; }

        string ClientSecret { get; set; }

        string CertificateThumbprint { get; set; }

        X509Certificate2 Certificate { get; set; }

        string RedirectUri { get; set; }

        /// <summary>
        /// EX China, Germany, USGovernment, Commercial
        /// </summary>
        AzureEnvironment AuthenticationEndpoint { get; set; }

        string TenantDomain { get; set; }

        string TenantId { get; set; }

        /// <summary>
        /// Represents the Azure AD Group claim to which the system should be locked down
        /// </summary>
        string SecurityGroupId { get; set; }

        /// <summary>
        /// Indicates if this should use Certificate Authentication
        /// </summary>
        /// <returns></returns>
        bool IsCertificateAuth();
    }
}
