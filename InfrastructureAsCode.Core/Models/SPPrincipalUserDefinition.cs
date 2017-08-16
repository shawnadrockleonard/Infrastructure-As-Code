using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Models
{
    public class SPPrincipalUserDefinition : SPPrincipalModel
    {
        public string Email { get; set; }
    }
}
