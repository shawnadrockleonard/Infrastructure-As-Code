using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Models
{
    /// <summary>
    /// Represents a threshold parsing CAML query
    /// </summary>
    public class SPThresholdEnumerationModel
    {
        /// <summary>
        /// Empty for LINQ
        /// </summary>
        public SPThresholdEnumerationModel() { }

        /// <summary>
        /// The GEG ID
        /// </summary>
        public int StartsWithId { get; set; }

        /// <summary>
        /// The LEG ID
        /// </summary>
        public int EndsWithId { get; set; }

        /// <summary>
        /// The resulting AND clause for the ID combo
        /// </summary>
        public string AndClause
        {
            get
            {
                return CAML.And(
                            CAML.Geq(CAML.FieldValue("ID", FieldType.Integer.ToString("f"), StartsWithId.ToString())),
                            CAML.Leq(CAML.FieldValue("ID", FieldType.Integer.ToString("f"), EndsWithId.ToString())));
            }
        }
    }
}
