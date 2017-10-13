using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Models
{
    public class SPOTaxonomyModel
    {
        public SPOTaxonomyModel()
        {
            TermStore = new SPOTaxonomyTermStoreModel()
            {
                Id = Guid.Empty
            };
            TermSet = new SPOTaxonomyTermSetModel()
            {
                Id = Guid.Empty
            };

            Customization = new List<KeyValuePair<string, string>>();
        }

        public string TermSetName { get; set; }

        public SPOTaxonomyTermStoreModel TermStore { get; set; }

        public SPOTaxonomyTermSetModel TermSet { get; set; }

        public List<KeyValuePair<string, string>> Customization { get; set; }
    }
}
