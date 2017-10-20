using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Models.NativeCSOM
{
    public class NativeFieldUserValue : NativeFieldLookupValue
    {
        public NativeFieldUserValue() : base() { }

        public string Email { get; set; }
    }
}
