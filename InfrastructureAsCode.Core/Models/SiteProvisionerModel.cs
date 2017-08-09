using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace InfrastructureAsCode.Core.Models
{
    /// <summary>
    /// Represents a JSON file for provisioning lists, groups, pages, views, etc
    /// </summary>
    public class SiteProvisionerModel
    {
        /// <summary>
        /// initialize collections
        /// </summary>
        public SiteProvisionerModel()
        {
        }

        /// <summary>
        /// Represents the [Me] component of a CAML query
        /// </summary>
        public string UserIdFilter { get { return "<UserID Type='Integer'/>"; } }

        /// <summary>
        /// Contains the web scoped content types
        /// </summary>
        public List<SPContentTypeDefinition> ContentTypes { get; set; }

        /// <summary>
        /// Contains the web scoped definitions
        /// </summary>
        public List<SPFieldDefinitionModel> FieldDefinitions { get; set; }

        /// <summary>
        /// Collection of List definitions
        /// </summary>
        public List<SPListDefinition> Lists { get; set; }

        /// <summary>
        /// Collection of groups for the Web/Site
        /// </summary>
        public List<SPGroupDefinitionModel> Groups { get; set; }

        /// <summary>
        /// Contains a collection of choices for internal names
        /// </summary>
        public List<SiteProvisionerFieldChoiceModel> FieldChoices { get; set; }
    }
}