using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace IaC.Core.Models
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
            this.Lists = new List<SPListDefinition>();
            this.Groups = new List<SPGroupDefinitionModel>();
            this.SiteResources = false;
        }

        public bool SiteResources { get; set; }

        /// <summary>
        /// Collection of List definitions
        /// </summary>
        public List<SPListDefinition> Lists { get; set; }

        /// <summary>
        /// Collection of groups for the Web/Site
        /// </summary>
        public List<SPGroupDefinitionModel> Groups { get; set; }
    }
}