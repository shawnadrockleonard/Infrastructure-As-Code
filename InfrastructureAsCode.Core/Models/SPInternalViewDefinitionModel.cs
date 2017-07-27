using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Models
{
    public class SPInternalViewDefinitionModel : SPViewDefinitionModel
    {
        public SPInternalViewDefinitionModel() : base()
        {

        }

        /// <summary>
        /// Represents the relative URL from the Web URL; Do not include a preceeding slash
        /// </summary>
        public string SitePage { get; set; }
    }
}
