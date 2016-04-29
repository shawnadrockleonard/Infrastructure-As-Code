using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;

namespace IaC.Core.Models
{
    public class SPWebDefinitionModel
    {
        public SPWebDefinitionModel()
        {
            this.Lists = new List<SPListDefinition>();
        }

        public DateTime Created { get; set; }
        public DateTime LastItemModifiedDate { get; set; }
        public int ListCount { get; set; }
        public int ListItemCount { get; set; }
        public IList<SPListDefinition> Lists { get; set; }
        public string SiteOwner { get; set; }
        public string SiteUrl { get; set; }
        public int UIVersion { get; set; }
        public UsageInfo UsageInfo { get; set; }
    }
}
