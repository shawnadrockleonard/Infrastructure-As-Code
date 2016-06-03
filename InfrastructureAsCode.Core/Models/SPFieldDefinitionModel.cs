using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Entities;
using System.Collections.Generic;
using System.Linq;

namespace InfrastructureAsCode.Core.Models
{
    /// <summary>
    /// Field Definition concrete class
    /// </summary>
    public class SPFieldDefinitionModel
    {
        public SPFieldDefinitionModel()
        {
            this.FieldGuid = new System.Guid();
            this.RestrictedMode = true;
            this.GroupName = "CustomDevelopment";
            this.RichTextField = false;
            this.HiddenField = false;
            this.AddToDefaultView = false;
            this.NumLines = 0;
            this.ChoiceFormat = ChoiceFormatType.Dropdown;
            this.DateFieldFormat = DateTimeFieldFormatType.DateTime;
            this.FieldChoices = new List<SPChoiceModel>();
        }

        public SPFieldDefinitionModel(FieldType fType) : this()
        {
            this.fieldType = fType;
        }

        /// <summary>
        /// Unique Identifier
        /// </summary>
        public System.Guid FieldGuid { get; set; }

        /// <summary>
        /// The custom group name in which this field will display
        /// </summary>
        public string GroupName { get; set; }

        /// <summary>
        /// Internal name for the field definition
        /// </summary>
        public string InternalName { get; set; }

        /// <summary>
        /// Converts an SP field with property internal name structure
        /// </summary>
        public string DisplayNameMasked
        {
            get
            {
                if (!string.IsNullOrEmpty(DisplayName))
                {
                    return this.DisplayName.Replace(" ", "_x0020_");
                }
                return InternalName;
            }
        }

        public string DisplayName { get; set; }

        public FieldType fieldType { get; set; }

        public bool AddToDefaultView { get; set; }

        public bool HiddenField { get; set; }

        public string Description { get; set; }

        public int NumLines { get; set; }

        public bool RichTextField { get; set; }

        public bool RestrictedMode { get; set; }

        public bool AppendOnly { get; set; }

        public DateTimeFieldFormatType? DateFieldFormat { get; set; }

        public bool FieldIndexed { get; set; }

        public List<SPChoiceModel> FieldChoices { get; set; }

        public string ChoiceDefault
        {
            get
            {
                if (FieldChoices.Count > 0)
                {
                    var sel = FieldChoices.FirstOrDefault(s => s.DefaultChoice);
                    if (sel != null)
                    {
                        return sel.Choice.Trim();
                    }
                }
                return string.Empty;
            }
        }

        public ChoiceFormatType ChoiceFormat { get; set; }

        public bool MultiChoice { get; set; }

        public string PeopleGroupName { get; set; }

        public bool PeopleOnly { get; set; }

        public string PeopleLookupField { get; set; }

        public int? MaxLength { get; set; }

        public bool Required { get; set; }

        /// <summary>
        /// If configure serialize the json file for choices
        /// </summary>
        public string LoadFromJSON { get; set; }

        /// <summary>
        /// Project the field defintion into the expected provisioning CSOM object
        /// </summary>
        /// <returns></returns>
        public FieldCreationInformation ToCreationObject()
        {
            var finfo = new FieldCreationInformation(this.fieldType);
            finfo.Id = this.FieldGuid;
            finfo.InternalName = this.InternalName;
            finfo.DisplayName = this.DisplayName;
            finfo.Group = this.GroupName;
            finfo.AddToDefaultView = this.AddToDefaultView;
            finfo.Required = this.Required;
            finfo.AdditionalAttributes = new List<KeyValuePair<string, string>>();

            return finfo;
        }
    }
}
