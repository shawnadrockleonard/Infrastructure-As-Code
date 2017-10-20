using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Models.NativeCSOM
{
    public class NativeFieldLookupValue
    {
        public int LookupId { get; set; }

        public string LookupValue { get; set; }

        public string TypeId { get; set; }
    }
}
