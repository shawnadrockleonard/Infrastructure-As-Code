using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Models
{
    public class SPGroupPrincipalModel : SPPrincipalModel
    {
        public SPGroupPrincipalModel() : base() { }

        public int GroupId { get; set; }

        public string GroupName { get; set; }

        public string GroupLogin { get; set; }

        public bool GroupHidden { get; set; }
    }
}

