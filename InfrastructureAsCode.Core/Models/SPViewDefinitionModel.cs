using InfrastructureAsCode.Core.Models.Enums;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace InfrastructureAsCode.Core.Models
{
    public class SPViewDefinitionModel
    {
        public SPViewDefinitionModel()
        {
            this.ViewQuery = null;
            this.Paged = true;
            this.FieldRefName = new List<string>();
            this.ToolBarType = ViewToolBarEnum.Standard;
        }

        public string Title { get; set; }

        public string InternalName { get; set; }

        public string CalculatedInternalName
        {
            get
            {
                if (!string.IsNullOrEmpty(this.InternalName))
                    return this.InternalName;

                return this.Title.Replace(" ", string.Empty);
            }
        }
        public Guid Id { get; set; }


        /// <summary>
        /// Represents the relative URL from the Web URL; Do not include a preceeding slash
        /// </summary>
        /// <remarks>Will only be populated if its internal view</remarks>
        public string SitePage { get; set; }

        public Nullable<bool> InternalView { get; set; }

        public string Aggregations { get; set; }

        public string AggregationsStatus { get; set; }

        public string BaseViewId { get; set; }

        public Nullable<bool> Hidden { get; set; }


        public string ImageUrl { get; set; }

        /// <summary>
        /// Collection of Fields referenced in the view
        /// </summary>
        public List<string> FieldRefName { get; set; }



        public bool DefaultView { get; set; }
        public bool MobileDefaultView { get; set; }
        public string ModerationType { get; set; }
        public bool OrderedView { get; set; }
        public ListPageRenderType PageRenderType { get; set; }

        public bool PersonalView { get; set; }
        
        public bool Paged { get; set; }
        public bool ReadOnlyView { get; set; }

        public uint RowLimit { get; set; }

        public ViewScope Scope { get; set; }
        public string StyleId { get; set; }
        public bool TabularView { get; set; }
        public bool Threaded { get; set; }
        public string Toolbar { get; set; }
        public string ToolbarTemplateName { get; }
        public Nullable<ViewToolBarEnum> ToolBarType { get; set; }

        public string ViewQuery { get; set; }
        
        public string ListViewXml { get; set; }

        public string ViewJoins { get; set; }

        public ViewType ViewCamlType { get; set; }

        public List<string> JsLinkFiles { get; set; }

        public bool HasJsLink
        {
            get
            {
                return (JsLinkFiles != null && JsLinkFiles.Count() > 0);
            }
        }

        public string JsLink
        {
            get
            {
                if (this.HasJsLink)
                {
                    return string.Join("|", this.JsLinkFiles);
                }
                return null;
            }
        }

    }
}