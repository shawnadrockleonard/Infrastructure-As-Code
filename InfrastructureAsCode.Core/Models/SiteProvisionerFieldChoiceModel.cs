using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Models
{
    /// <summary>
    /// Represents a JSON file for internal fields and their choices
    /// </summary>
    public class SiteProvisionerFieldChoiceModel
    {
        public SiteProvisionerFieldChoiceModel()
        {
            Choices = new List<SPChoiceModel>();
        }

        public string FieldInternalName { get; set; }

        public List<SPChoiceModel> Choices { get; set; }
    }
}
