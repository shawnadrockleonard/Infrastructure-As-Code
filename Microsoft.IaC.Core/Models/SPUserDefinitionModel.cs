using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.IaC.Core.Models;

namespace Microsoft.IaC.Core.Models
{
    public class SPUserDefinitionModel
    {
        public Guid Id { get; set; }

        public string UserName { get; set; }

        public string UserEmail { get; set; }

        public string UserDisplay { get; set; }

        public string Organization { get; set; }

        public string Manager { get; set; }
        public string OrganizationAcronym { get; set; }
    }
}
