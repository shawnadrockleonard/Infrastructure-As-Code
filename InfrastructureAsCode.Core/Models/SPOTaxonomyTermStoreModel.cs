using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Models
{
    /// <summary>
    /// Represents the Term Store
    /// </summary>
    public class SPOTaxonomyTermStoreModel
    {
        public Guid Id { get; set; }

        public string Name { get; set; }

        public bool IsOnline { get; set; }

        public int DefaultLanguage { get; set; }

        public string ContentTypePublishingHub { get; set; }

        public int WorkingLanguage { get; set; }
    }
}
