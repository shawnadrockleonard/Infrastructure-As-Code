using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Reports.o365Graph.TenantReport
{
    public class JSONAuditCollection<T> where T : class
    {
        /// <summary>
        /// initilize collections
        /// </summary>
        public JSONAuditCollection()
        {
            this.value = new List<T>();
        }

        /// <summary>
        /// Represents the metadata regarding the ODAta API service
        /// </summary>
        [JsonProperty("@odata.context")]
        public string Metadata { get; set; }

        /// <summary>
        /// Provides the next ODATA Paging link
        /// </summary>
        [JsonProperty("@odata.nextLink")]
        public string NextLink { get; set; }

        /// <summary>
        /// Serializable collection of auditiable events
        /// </summary>
        public List<T> value { get; set; }
    }
}
