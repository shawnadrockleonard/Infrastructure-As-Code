using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Reports.o365Graph.AzureAD
{
    /// <summary>
    /// Azure AD Config settings for claiming tokens
    /// </summary>
    public interface IAzureADConfig
    {

        string CallbackPath { get; set; }

        string ClientId { get; set; }

        string ClientSecret { get; set; }

        string RedirectUri { get; set; }

        string TenantDomain { get; set; }

        string TenantId { get; set; }

        /// <summary>
        /// Represents the Azure AD Group claim to which the system should be locked down
        /// </summary>
        string SecurityGroupId { get; set; }
    }
}
