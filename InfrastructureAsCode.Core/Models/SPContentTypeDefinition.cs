using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Models
{
    public class SPContentTypeDefinition
    {
        /// <summary>
        /// initialize arrays
        /// </summary>
        public SPContentTypeDefinition()
        {
            this.FieldLinks = new List<SPFieldLinkDefinitionModel>();
            this.ContentTypeGroup = "CustomDevelopment";
        }

        public string ContentTypeId { get; set; }

        public string Name { get; set; }

        public string Description { get; set; }
        
        public bool Inherits { get; set; }

        public bool DefaultContentType { get; set; }

        public List<SPFieldLinkDefinitionModel> FieldLinks { get; set; }

        public string ContentTypeGroup { get; set; }

        public string DocumentTemplate { get; set; }

        public bool Hidden { get; set; }

        public string JSLink { get; set; }

        public string Scope { get; set; }

        public ContentTypeCreationInformation ToCreationObject()
        {
            var info = new ContentTypeCreationInformation()
            {
                Id = this.ContentTypeId,
                Name = this.Name,
                Description = this.Description,
                Group = this.ContentTypeGroup
            };
            return info;
        }
    }
}
