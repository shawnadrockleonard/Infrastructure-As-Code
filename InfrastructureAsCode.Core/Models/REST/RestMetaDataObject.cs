using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Models.REST
{
    public class RestMetaDataObject
    {
        public Guid id { get; set; }

        public string uri { get; set; }

        public string etag { get; set; }

        [JsonProperty("type")]
        public string resttype { get; set; }
    }
}
