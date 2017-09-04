using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Models.REST
{
    public class RestListItemResponseMinimalObject<T> where T : IRestListItemObj
    {
        [JsonProperty("odata.metadata")]
        public string Metadata { get; set; }

        [JsonProperty("odata.nextLink")]
        public string NextLink { get; set; }

        public List<T> value { get; set; }
    }
}
