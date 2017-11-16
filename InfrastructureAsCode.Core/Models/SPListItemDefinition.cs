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

        /// <summary>
        /// The title column for the listitem
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// The date the item was created
        /// </summary>
        public Nullable<DateTime> Created { get; set; }

        /// <summary>
        /// User who created the item
        /// </summary>
        public SPPrincipalUserDefinition CreatedBy { get; set; }

        /// <summary>
        /// The date the item was last modified
        /// </summary>
        public Nullable<DateTime> Modified { get; set; }

        /// <summary>
        /// User who last modified the item
        /// </summary>
        public SPPrincipalUserDefinition ModifiedBy { get; set; }

        /// <summary>
        /// The URL to the Item
        /// </summary>
        public string FileRef { get; set; }

        /// <summary>
        /// The URL to the folder containing the Item
        /// </summary>
        public string FileDirRef { get; set; }

        /// <summary>
        /// Contains the ListItem Column data
        /// </summary>
        public List<SPListItemFieldDefinition> ColumnValues { get; set; }

        /// <summary>
        /// A collection of specialized roles
        /// </summary>
        public IList<SPPrincipalModel> RoleBindings { get; set; }
    }
}
