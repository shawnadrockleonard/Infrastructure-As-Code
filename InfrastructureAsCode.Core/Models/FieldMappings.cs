using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Models
{

    public class FieldMappings
    {
        public Microsoft.SharePoint.Client.FieldType ColumnType { get; set; }

        public string ColumnInternalName { get; set; }

        public bool ColumnMandatory { get; set; }
    }
}
