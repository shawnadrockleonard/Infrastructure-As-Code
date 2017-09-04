using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Models.REST
{
    public class RestBaseObject
    {
        public RestMetaDataObject __metadata { get; set; }

        public RestDeferredObj RoleAssignments { get; set; }
    }
}
