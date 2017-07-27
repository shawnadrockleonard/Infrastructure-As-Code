using Microsoft.SharePoint.Client.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Models
{
    public class SPPrincipalModel
    {
        public int Id { get; set; }

        public bool IsHiddenInUI { get; set; }

        public string LoginName { get; set; }

        public PrincipalType PrincipalType { get; set; }

        public string Title { get; set; }
    }
}
