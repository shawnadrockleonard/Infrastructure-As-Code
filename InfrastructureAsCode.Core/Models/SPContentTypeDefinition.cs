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
            this.FieldLinkRefs = new List<string>();
            this.ContentTypeGroup = "CustomDevelopment";
        }

        public string ContentTypeId { get; set; }

        public string ContentTypeName { get; set; }

        public string ContentTypeDescription { get; set; }
        
        public bool Inherits { get; set; }

        public bool DefaultContentType { get; set; }
        
        public List<string> FieldLinkRefs { get; set; }

        public string ContentTypeGroup { get; set; }

        public ContentTypeCreationInformation ToCreationObject()
        {
            var info = new ContentTypeCreationInformation()
            {
                Id = this.ContentTypeId,
                Name = this.ContentTypeName,
                Description = this.ContentTypeDescription,
                Group = this.ContentTypeGroup
            };
            return info;
        }
    }
}
