using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.IaC.Core.Models
{
    public class SPListItemDefinition
    {
        public string Title { get; set; }

        public List<SPListItemFieldDefinition> ColumnValues { get; set; }

        public SPListItemDefinition()
        {
            this.ColumnValues = new List<SPListItemFieldDefinition>();
        }
    }
}
