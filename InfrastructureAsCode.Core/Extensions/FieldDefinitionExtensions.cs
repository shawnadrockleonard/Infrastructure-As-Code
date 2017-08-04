using InfrastructureAsCode.Core.Models;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace InfrastructureAsCode.Core.Extensions
{
    /// <summary>
    /// Provides extensions class for extending OfficeDevPnP to provision SP Fields
    /// </summary>
    public static class FieldDefinitionExtensions
    {
        /// <summary>
        /// Builds a Field Creation Info object from the Definition model and returns its resulting XML
        /// </summary>
        /// <param name="host">The instantiated web/list/library to which the field will be added</param>
        /// <param name="fieldDefinition">The definition pulled from a SP Site or user construction</param>
        /// <param name="siteGroups">The collection of site groups that a user/group field make filter</param>
        /// <param name="JsonFilePath">The file path to a JSON file containing Choice lookups</param>
        /// <returns></returns>
        public static string CreateFieldDefinition(this SecurableObject host, SPFieldDefinitionModel fieldDefinition, List<SPGroupDefinitionModel> siteGroups, string JsonFilePath = null)
        {
            var idguid = fieldDefinition.FieldGuid;
            var choiceXml = string.Empty;
            var defaultChoiceXml = string.Empty;
            var attributes = new List<KeyValuePair<string, string>>();

            if (string.IsNullOrEmpty(fieldDefinition.InternalName))
            {
                throw new ArgumentNullException("InternalName");
            }

            if (string.IsNullOrEmpty(fieldDefinition.Title))
            {
                throw new ArgumentNullException("DisplayName");
            }

            if (!string.IsNullOrEmpty(fieldDefinition.LoadFromJSON) && string.IsNullOrEmpty(JsonFilePath))
            {
                throw new ArgumentNullException("JsonFilePath", "You must specify a file path to the JSON file if loading from JSON");
            }

            if (!string.IsNullOrEmpty(fieldDefinition.PeopleGroupName) && (siteGroups == null || siteGroups.Count() <= 0))
            {
                throw new ArgumentNullException("SiteGroups", string.Format("You must specify a collection of group for the field {0}", fieldDefinition.Title));
            }

            if (string.IsNullOrEmpty(fieldDefinition.LookupListName) && fieldDefinition.FieldTypeKind == FieldType.Lookup)
            {
                throw new ArgumentNullException("LookupListName", string.Format("you must specify a lookup list title for the field {0}", fieldDefinition.Title));
            }


            if (!string.IsNullOrEmpty(fieldDefinition.Description))
            {
                attributes.Add(new KeyValuePair<string, string>("Description", fieldDefinition.Description));
            }
            if (fieldDefinition.FieldIndexed)
            {
                attributes.Add(new KeyValuePair<string, string>("Indexed", fieldDefinition.FieldIndexed.ToString().ToUpper()));
            }
            if(fieldDefinition.AppendOnly)
            {
                attributes.Add(new KeyValuePair<string, string>("AppendOnly", fieldDefinition.AppendOnly.ToString().ToUpper()));
            }

            var choices = new FieldType[] { FieldType.Choice, FieldType.GridChoice, FieldType.MultiChoice, FieldType.OutcomeChoice };
            if (choices.Any(a => a == fieldDefinition.FieldTypeKind))
            {
                if (!string.IsNullOrEmpty(fieldDefinition.LoadFromJSON))
                {
                    var filePath = string.Format("{0}\\{1}", JsonFilePath, fieldDefinition.LoadFromJSON);
                    //#TODO: Check file path
                    var contents = JsonConvert.DeserializeObject<List<SPChoiceModel>>(System.IO.File.ReadAllText(filePath));
                    fieldDefinition.FieldChoices.Clear();
                    fieldDefinition.FieldChoices.AddRange(contents);
                }

                choiceXml = string.Format("<CHOICES>{0}</CHOICES>", string.Join("", fieldDefinition.FieldChoices.Select(s => string.Format("<CHOICE>{0}</CHOICE>", s.Choice.Trim())).ToArray()));
                if (!string.IsNullOrEmpty(fieldDefinition.ChoiceDefault))
                {
                    defaultChoiceXml = string.Format("<Default>{0}</Default>", fieldDefinition.ChoiceDefault);
                }
                if (fieldDefinition.FieldTypeKind == FieldType.Choice)
                {
                    attributes.Add(new KeyValuePair<string, string>("Format", fieldDefinition.ChoiceFormat.ToString("f")));
                }

            }
            else if (fieldDefinition.FieldTypeKind == FieldType.DateTime)
            {
                if (fieldDefinition.DateFieldFormat.HasValue)
                {
                    attributes.Add(new KeyValuePair<string, string>("DisplayFormat", fieldDefinition.DateFieldFormat.Value.ToString("f")));
                }
            }
            else if (fieldDefinition.FieldTypeKind == FieldType.Note)
            {
                attributes.Add(new KeyValuePair<string, string>("RichText", fieldDefinition.RichTextField.ToString().ToUpper()));
                attributes.Add(new KeyValuePair<string, string>("NumLines", fieldDefinition.NumLines.ToString()));
                if (fieldDefinition.RestrictedMode)
                {
                    attributes.Add(new KeyValuePair<string, string>("RestrictedMode", fieldDefinition.RestrictedMode.ToString().ToUpper()));
                }
                else
                {
                    attributes.Add(new KeyValuePair<string, string>("RichTextMode", "FullHtml"));
                    attributes.Add(new KeyValuePair<string, string>("IsolateStyles", true.ToString().ToUpper()));
                }

            }
            else if (fieldDefinition.FieldTypeKind == FieldType.User)
            {
                //AllowMultipleValues
                if (fieldDefinition.MultiChoice)
                {
                    attributes.Add(new KeyValuePair<string, string>("Mult", fieldDefinition.MultiChoice.ToString().ToUpper()));
                }
                //SelectionMode
                if (fieldDefinition.PeopleOnly)
                {
                    attributes.Add(new KeyValuePair<string, string>("UserSelectionMode", FieldUserSelectionMode.PeopleOnly.ToString("d")));
                }

                if (!string.IsNullOrEmpty(fieldDefinition.PeopleLookupField))
                {
                    attributes.Add(new KeyValuePair<string, string>("ShowField", fieldDefinition.PeopleLookupField));
                    //fldUser.LookupField = fieldDef.PeopleLookupField;
                }
                if (!string.IsNullOrEmpty(fieldDefinition.PeopleGroupName))
                {
                    var group = siteGroups.FirstOrDefault(f => f.Title == fieldDefinition.PeopleGroupName);
                    if (group != null)
                    {
                        attributes.Add(new KeyValuePair<string, string>("UserSelectionScope", group.Id.ToString()));
                    }
                }
            }
            else if (fieldDefinition.FieldTypeKind == FieldType.Lookup)
            {
                var lParentList = host.GetAssociatedWeb().GetListByTitle(fieldDefinition.LookupListName);
                var strParentListID = lParentList.Id;

                attributes.Add(new KeyValuePair<string, string>("EnforceUniqueValues", false.ToString().ToUpper()));
                attributes.Add(new KeyValuePair<string, string>("List", strParentListID.ToString("B")));
                attributes.Add(new KeyValuePair<string, string>("ShowField", fieldDefinition.LookupListFieldName));
            }

            var finfo = fieldDefinition.ToCreationObject();
            finfo.AdditionalAttributes = attributes;
            var finfoXml = FieldAndContentTypeExtensions.FormatFieldXml(finfo);
            if (!string.IsNullOrEmpty(choiceXml))
            {
                XDocument xd = XDocument.Parse(finfoXml);
                XElement root = xd.FirstNode as XElement;
                if (!string.IsNullOrEmpty(defaultChoiceXml))
                {
                    root.Add(XElement.Parse(defaultChoiceXml));
                }
                root.Add(XElement.Parse(choiceXml));
                finfoXml = xd.ToString();
            }

            return finfoXml;
        }
    }
}
