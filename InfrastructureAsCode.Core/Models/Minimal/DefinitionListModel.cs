using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Models.Minimal
{
    public class DefinitionListModel
    {
        public DefinitionListModel()
        {
            this.fields = new List<DefinitionFieldModel>();
        }

        public Guid id { get; set; }

        public string title { get; set; }

        public string url { get; set; }

        public bool contentTypes { get; set; }

        public List<DefinitionFieldModel> fields { get; set; }
    }
}
