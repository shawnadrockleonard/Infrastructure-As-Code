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
            this.QueryXml = null;
            this.PagedView = true;
            this.FieldRefName = new string[] { };
            this.JsLinkFiles = new string[] { };
            this.ToolBarType = ViewToolBarEnum.Standard;
        }

        public string Title { get; set; }

        public string InternalName
        {
            get
            {
                return this.Title.Replace(" ", string.Empty);
            }
        }

        public uint RowLimit { get; set; }

        public string[] FieldRefName { get; set; }

        public ViewType ViewCamlType { get; set; }

        public ViewToolBarEnum ToolBarType { get; set; }

        public bool DefaultView { get; set; }

        public string QueryXml { get; set; }

        public string[] JsLinkFiles { get; set; }

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

        public bool PersonalView { get; set; }

        public bool PagedView { get; set; }
    }
}