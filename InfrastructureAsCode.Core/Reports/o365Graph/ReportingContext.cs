using InfrastructureAsCode.Core.Reports.o365Graph.AzureAD;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Reports.o365Graph
{
    /// <summary>
    /// MS Graph Reporting endpoints currently in beta
    /// </summary>
    public class ReportingContext
    {
        #region Privates

        /// <summary>
        /// Represents the Graph API endpoints
        /// </summary>
        /// <remarks>Of note this is a BETA inpoint as these APIs are in Preview</remarks>
        internal static string defaultServiceEndpointUrl = "https://graph.microsoft.com/beta/reports/{0}({1})/content";

        #endregion Privates

        #region Properties

        public string GraphUrl
        {
            get;
            set;
        }

        #endregion

        public ReportingContext()
            : this(defaultServiceEndpointUrl)
        {
        }

        public ReportingContext(string url)
        {
            this.GraphUrl = url;
        }


    }
}
