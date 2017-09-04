using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Models.REST
{
    public class RestListItemResponseVerboseObject<T> where T : IRestListItemObj
    {
        public List<T> results { get; set; }

        public string __next { get; set; }

    }
}
