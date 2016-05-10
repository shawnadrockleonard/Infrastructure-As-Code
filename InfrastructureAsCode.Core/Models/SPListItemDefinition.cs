using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Models
{
    public class SPListItemDefinition
    {
        /// <summary>
        /// Represents the List Item unique ID, only necessary for reporting and query match
        /// </summary>
        public int? ID { get; set; }

        /// <summary>
        /// Represents the List Item guid, only necessary for reporting and query match
        /// </summary>
        public Guid? ItemID { get; set; }

        public string Title { get; set; }

        public List<SPListItemFieldDefinition> ColumnValues { get; set; }

        public SPListItemDefinition()
        {
            this.ColumnValues = new List<SPListItemFieldDefinition>();
        }
    }
}
