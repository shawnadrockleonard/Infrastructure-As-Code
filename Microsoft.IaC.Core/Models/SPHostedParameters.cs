using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace IaC.Core.Models
{
    /// <summary>
    /// SharePoint Querystring parameters for serialization
    /// </summary>
    public class SPHostedParameters
    {
        /// <summary>
        /// The host web for the add-in
        /// </summary>
        public string SPHostUrl { get; set; }

        /// <summary>
        /// The language for translation
        /// </summary>
        public string SPLanguage { get; set; }

        /// <summary>
        /// Client running version
        /// </summary>
        public int SPClientTag { get; set; }

        /// <summary>
        /// Product running version
        /// </summary>
        public string SPProductNumber { get; set; }

        /// <summary>
        /// The appweb for the add-in
        /// </summary>
        public string SPAppWebUrl { get; set; }
    }
}