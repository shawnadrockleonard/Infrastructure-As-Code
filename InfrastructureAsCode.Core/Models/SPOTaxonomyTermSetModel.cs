using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Models
{
    /// <summary>
    /// Represents a Term Set within a Term Store
    /// </summary>
    public class SPOTaxonomyTermSetModel : SPOTaxonomyItemModel
    {
        public string CustomSortOrder { get; set; }

        public bool IsAvailableForTagging { get; set; }

        public string Owner { get; set; }

        public string Description { get; set; }

        public Guid TermStoreId { get; set; }

        public bool IsOpenForTermCreation { get; set; }

        public SPOTaxonomyItemModel Group { get; set; }
    }
}
