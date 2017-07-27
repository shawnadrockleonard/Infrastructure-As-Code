using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Models
{
    public class SPListItemDefinition : SPId
    {
        public SPListItemDefinition() : base()
        {
            this.ColumnValues = new List<SPListItemFieldDefinition>();
            this.RoleBindings = new List<SPPrincipalModel>();
        }

        public string Title { get; set; }

        public List<SPListItemFieldDefinition> ColumnValues { get; set; }

        public string ListDescription { get; set; }

        /// <summary>
        /// A collection of specialized roles
        /// </summary>
        public IList<SPPrincipalModel> RoleBindings { get; set; }
    }
}
