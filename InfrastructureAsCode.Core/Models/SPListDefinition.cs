using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Models
{
    /// <summary>
    /// Defines a SharePoint List or Library
    /// </summary>
    public class SPListDefinition
    {
        public SPListDefinition()
        {
            this.ListTemplate = ListTemplateType.GenericList;
            this.ContentTypeEnabledOverride = false;
            this.Views = new List<SPViewDefinitionModel>();
            this.ContentTypes = new List<SPContentTypeDefinition>();
            this.ListItems = new List<SPListItemDefinition>();
            this.FieldDefinitions = new List<SPFieldDefinitionModel>();
        }


        public string ListName { get; set; }


        public string InternalName
        {
            get
            {
                return this.ListName.Replace(" ", string.Empty);
            }
        }

        public string ListDescription { get; set; }

        public QuickLaunchOptions QuickLaunch { get; set; }

        public ListTemplateType ListTemplate { get; set; }

        /// <summary>
        /// List/Library should have content types enabled but no content types are specified
        /// </summary>
        public bool ContentTypeEnabledOverride { get; set; }

        /// <summary>
        /// If content types are specified or override is set [Enable] content types
        /// </summary>
        public bool ContentTypeEnabled
        {
            get
            {
                if (ContentTypeEnabledOverride) return true;
                return HasContentTypes;
            }
        }

        public List<SPViewDefinitionModel> Views { get; set; }

        public List<SPContentTypeDefinition> ContentTypes { get; set; }

        /// <summary>
        /// If content types are specified or override is set [Enable] content types
        /// </summary>
        public bool HasContentTypes
        {
            get
            {
                if (ContentTypes != null && ContentTypes.Count > 0)
                    return true;
                return false;
            }
        }
        public List<SPFieldDefinitionModel> FieldDefinitions { get; set; }

        public List<SPListItemDefinition> ListItems { get; set; }

        /// <summary>
        /// Should the list/library allow versioning
        /// </summary>
        public bool Versioning { get; set; }
        public bool EnableFolderCreation { get; set; }

        /// <summary>
        /// Projects the properties to Creation Info
        /// </summary>
        /// <returns></returns>
        public ListCreationInformation ToCreationObject()
        {
            var info = new ListCreationInformation()
            {
                Title = this.ListName,
                Description = this.ListDescription,
                QuickLaunchOption = this.QuickLaunch,
                TemplateType = (int)this.ListTemplate,
                Url = this.InternalName
            };
            return info;
        }
    }
}
