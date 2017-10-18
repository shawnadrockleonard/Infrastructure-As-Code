using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace InfrastructureAsCode.Core.Constants
{
    /// <summary>
    /// Contains XML namespaces for processing various SharePoint XML
    /// </summary>
    public static class ConstantsXmlNamespaces
    {
        /// <summary>
        /// Xml Generic Schema
        /// </summary>
        public static XNamespace ListNS = "http://www.w3.org/2001/XMLSchema";

        /// <summary>
        /// Xml Generic instance schema
        /// </summary>
        public static XNamespace InstanceNS = "http://www.w3.org/2001/XMLSchema-instance";

        /// <summary>
        /// SharePoint URI for XML parsing
        /// </summary>
        public static XNamespace SharePointNS = "http://schemas.microsoft.com/sharepoint/";
    }
}
