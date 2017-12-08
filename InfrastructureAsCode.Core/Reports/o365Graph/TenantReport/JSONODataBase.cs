using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Reports.o365Graph.TenantReport
{
    public class JSONODataBase
    {
        /// <summary>
        /// Represents the OData API data type
        /// </summary>
        [JsonProperty("@odata.type")]
        public string ODataType { get; set; }
    }
}
