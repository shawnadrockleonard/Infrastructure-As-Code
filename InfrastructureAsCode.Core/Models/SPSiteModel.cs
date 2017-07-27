using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Models
{
    /// <summary>
    /// A Site definition for SharePoint
    /// </summary>
    public class SPSiteModel : SPUniqueId
    {
        public SPSiteModel() : base()
        {
            this.Owners = new List<SPPrincipalModel>();
            this.Groups = new List<SPGroupPrincipalModel>();
            this.Users = new List<SPPrincipalModel>();
            this.Lists = new List<SPListDefinition>();
            this.ContentTypes = new List<SPContentTypeDefinition>();
            this.FieldDefinitions = new List<SPFieldDefinitionModel>();
        }

        /// <summary>
        /// The absolute URL for the Site Collection
        /// </summary>
        public string Url { get; set; }

        /// <summary>
        /// The title for the SharePoint site
        /// </summary>
        public string title { get; set; }

        /// <summary>
        /// Collection of owners for the site
        /// </summary>
        public IList<SPPrincipalModel> Owners { get; set; }

        /// <summary>
        /// Collection of groups for the site
        /// </summary>
        public List<SPGroupPrincipalModel> Groups { get; set; }

        /// <summary>
        /// Collection of users for the site
        /// </summary>
        public List<SPPrincipalModel> Users { get; set; }

        /// <summary>
        /// Collection of lists for the Site
        /// </summary>
        public List<SPListDefinition> Lists { get; set; }

        /// <summary>
        /// Collection of content types for the site
        /// </summary>
        public List<SPContentTypeDefinition> ContentTypes { get; set; }

        /// <summary>
        /// Collection of fields for the Site
        /// </summary>
        public List<SPFieldDefinitionModel> FieldDefinitions { get; set; }
    }
}
