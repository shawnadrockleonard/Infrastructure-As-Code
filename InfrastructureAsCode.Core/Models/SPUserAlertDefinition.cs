using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Models
{
    public class SPUserAlertDefinition
    {
        public string AlertForTitle { get; set; }
        public string AlertForUrl { get; set; }
        public string CurrentUser { get; set; }
        public string EventType { get; set; }
        public string Id { get; set; }
        public string WebId { get; set; }
        public string WebTitle { get; set; }
    }
}
