using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Models
{

    public class FeatureDefinition
    {
        public string DisplayName { get; set; }

        public Guid Id { get; set; }

        public FeatureDefinitionScope Scope { get; set; }

        public bool IsActivated { get; set; }
    }
}
