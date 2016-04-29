using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IaC.Core.Models
{
    public class SPGroupMSOnlineDefinition : SPGroupDefinitionModel
    {
        public SPGroupMSOnlineDefinition() : base()
        {

        }

        public string EmailAddress { get; set; }
        public string GroupType { get; set; }
        public string IsSystem { get; set; }
        public string LastDirSyncTime { get; set; }
        public string ManagedBy { get; set; }
        public string ObjectId { get; set; }
        public string ValidationStatus { get; set; }
    }
}
