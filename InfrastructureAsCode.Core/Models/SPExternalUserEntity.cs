using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core.Entities;

namespace InfrastructureAsCode.Core.Models
{
    public class SPExternalUserEntity : ExternalUserEntity
    {
        public Nullable<int> UserId { get; set; }

        public bool FoundInSiteUsers { get; set; }
    }
}
