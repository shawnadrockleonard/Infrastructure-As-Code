using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Reports.o365Graph.AzureAD
{
    /// <summary>
    /// Represents the config for claiming and refreshing tokens
    /// </summary>
    public class AzureADConfig : IAzureADConfig
    {
        public string CallbackPath { get; set; }

        public string ClientId { get; set; }

        public string ClientSecret { get; set; }

        public string RedirectUri { get; set; }

        public string TenantDomain { get; set; }

        public string TenantId { get; set; }

        public string SecurityGroupId { get; set; }
    }
}
