using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Models
{
    public class SPOTaxonomyItemModel
    {
        public SPOTaxonomyItemModel()
        {
            Id = Guid.Empty;
        }

        public Guid Id { get; set; }

        public string Name { get; set; }

        public DateTime CreatedDate { get; set; }

        public DateTime LastModifiedDate { get; set; }
    }
}
