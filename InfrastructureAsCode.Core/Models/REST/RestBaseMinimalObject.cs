using Newtonsoft.Json;
using System;

namespace InfrastructureAsCode.Core.Models.REST
{
    public class RestBaseMinimalObject
    {
        [JsonProperty("odata.id")]
        public Guid odataid { get; set; }

        [JsonProperty("odata.editLink")]
        public string odatauri { get; set; }

        [JsonProperty("odata.etag")]
        public string odataetag { get; set; }

        [JsonProperty("odata.type")]
        public string odatatype { get; set; }
    }
}