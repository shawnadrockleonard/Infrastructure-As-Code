using InfrastructureAsCode.Core.Models;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using OfficeDevPnP.Core.Utilities;
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
        /// <param name="provisionerChoices">(OPTIONAL) The collection of hoice lookups as defined in the serialized JSON file</param>
        /// <returns></returns>
        public static string CreateFieldDefinition(this SecurableObject host, SPFieldDefinitionModel fieldDefinition, List<SPGroupDefinitionModel> siteGroups, List<SiteProvisionerFieldChoiceModel> provisionerChoices = null)
        {
            var idguid = fieldDefinition.FieldGuid;
            var choiceXml = string.Empty;
            var defaultChoiceXml = string.Empty;
            var formulaXml = string.Empty;
            var fieldReferenceXml = string.Empty;
            var attributes = new List<KeyValuePair<string, string>>();

            if (string.IsNullOrEmpty(fieldDefinition.InternalName))
            {
                throw new ArgumentNullException("InternalName");
            }

            if (string.IsNullOrEmpty(fieldDefinition.Title))
            {
                throw new ArgumentNullException("DisplayName");
            }

            if (!string.IsNullOrEmpty(fieldDefinition.PeopleGroupName) && (siteGroups == null || siteGroups.Count() <= 0))
            {
                throw new ArgumentNullException("SiteGroups", string.Format("You must specify a collection of group for the field {0}", fieldDefinition.Title));
            }

            if (string.IsNullOrEmpty(fieldDefinition.LookupListName) && fieldDefinition.FieldTypeKind == FieldType.Lookup)
            {
                throw new ArgumentNullException("LookupListName", string.Format("you must specify a lookup list title for the field {0}", fieldDefinition.Title));
            }

            if (fieldDefinition.LoadFromJSON && (provisionerChoices == null || !provisionerChoices.Any(pc => pc.FieldInternalName == fieldDefinition.InternalName)))
            {
                throw new ArgumentNullException("provisionerChoices", string.Format("You must specify a collection of field choices for the field {0}", fieldDefinition.Title));
            }


            if (!string.IsNullOrEmpty(fieldDefinition.Description))
            {
                attributes.Add(new KeyValuePair<string, string>("Description", fieldDefinition.Description));
            }
            if (fieldDefinition.FieldIndexed)
            {
                attributes.Add(new KeyValuePair<string, string>("Indexed", fieldDefinition.FieldIndexed.ToString().ToUpper()));
            }
            if (fieldDefinition.AppendOnly)
            {
                attributes.Add(new KeyValuePair<string, string>("AppendOnly", fieldDefinition.AppendOnly.ToString().ToUpper()));
            }

            var choices = new FieldType[] { FieldType.Choice, FieldType.GridChoice, FieldType.MultiChoice, FieldType.OutcomeChoice };
            if (choices.Any(a => a == fieldDefinition.FieldTypeKind))
            {
                if (fieldDefinition.LoadFromJSON
                    && (provisionerChoices != null && provisionerChoices.Any(fc => fc.FieldInternalName == fieldDefinition.InternalName)))
                {
                    var choicecontents = provisionerChoices.FirstOrDefault(fc => fc.FieldInternalName == fieldDefinition.InternalName);
                    fieldDefinition.FieldChoices.Clear();
                    fieldDefinition.FieldChoices.AddRange(choicecontents.Choices);
                }

                //AllowMultipleValues
                if (fieldDefinition.MultiChoice)
                {
                    attributes.Add(new KeyValuePair<string, string>("Mult", fieldDefinition.MultiChoice.ToString().ToUpper()));
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
                    attributes.Add(new KeyValuePair<string, string>("Format", fieldDefinition.DateFieldFormat.Value.ToString("f")));
                }
            }
            else if (fieldDefinition.FieldTypeKind == FieldType.Note)
            {
                attributes.Add(new KeyValuePair<string, string>("RichText", fieldDefinition.RichTextField.ToString().ToUpper()));
                attributes.Add(new KeyValuePair<string, string>("RestrictedMode", fieldDefinition.RestrictedMode.ToString().ToUpper()));
                attributes.Add(new KeyValuePair<string, string>("NumLines", fieldDefinition.NumLines.ToString()));
                if (!fieldDefinition.RestrictedMode)
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
                if (fieldDefinition.MultiChoice)
                {
                    attributes.Add(new KeyValuePair<string, string>("Mult", fieldDefinition.MultiChoice.ToString().ToUpper()));
                }
            }
            else if (fieldDefinition.FieldTypeKind == FieldType.Calculated)
            {
                attributes.Add(new KeyValuePair<string, string>("ResultType", fieldDefinition.OutputType.Value.ToString("f")));
                formulaXml = string.Format("<Formula>{0}</Formula>", fieldDefinition.DefaultFormula.UnescapeXml().EscapeXml());
                fieldReferenceXml = string.Format("<FieldRefs>{0}</FieldRefs>", string.Join("", fieldDefinition.FieldReferences.Select(s => CAML.FieldRef(s.Trim())).ToArray()));
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
            if (!string.IsNullOrEmpty(formulaXml))
            {
                XDocument xd = XDocument.Parse(finfoXml);
                XElement root = xd.FirstNode as XElement;
                root.Add(XElement.Parse(formulaXml));
                if (!string.IsNullOrEmpty(fieldReferenceXml))
                {
                    root.Add(XElement.Parse(fieldReferenceXml));
                }
                finfoXml = xd.ToString();
            }

            return finfoXml;
        }


        /// <summary>
        /// Parse the Field into the portable field definition object
        /// </summary>
        /// <param name="field">The field from which we are pulling a portable definition</param>
        /// <param name="fieldWeb">The web in which the field is provisioned</param>
        /// <param name="siteGroups">A collection of SharePoint groups</param>
        /// <param name="schemaXml">(OPTIONAL) the schema xml parsed into an XDocument</param>
        /// <returns></returns>
        public static SPFieldDefinitionModel RetrieveField(this Field field, Web fieldWeb, IEnumerable<Group> siteGroups = null, XElement schemaXml = null)
        {
            field.EnsureProperties(
                lft => lft.Id,
                lft => lft.AutoIndexed,
                lft => lft.CanBeDeleted,
                lft => lft.DefaultFormula,
                lft => lft.DefaultValue,
                lft => lft.Group,
                lft => lft.Description,
                lft => lft.EnforceUniqueValues,
                lft => lft.FieldTypeKind,
                lft => lft.Filterable,
                lft => lft.FromBaseType,
                lft => lft.Hidden,
                lft => lft.Indexed,
                lft => lft.InternalName,
                lft => lft.JSLink,
                lft => lft.NoCrawl,
                lft => lft.ReadOnlyField,
                lft => lft.Required,
                lft => lft.SchemaXml,
                lft => lft.Scope,
                lft => lft.Title);

            var fieldModel = new SPFieldDefinitionModel()
            {
                FieldGuid = field.Id,
                AutoIndexed = field.AutoIndexed,
                CanBeDeleted = field.CanBeDeleted,
                DefaultFormula = field.DefaultFormula,
                DefaultValue = field.DefaultValue,
                GroupName = field.Group,
                Description = field.Description,
                EnforceUniqueValues = field.EnforceUniqueValues,
                FieldTypeKind = field.FieldTypeKind,
                Filterable = field.Filterable,
                HiddenField = field.Hidden,
                FieldIndexed = field.Indexed,
                InternalName = field.InternalName,
                JSLink = field.JSLink,
                NoCrawl = field.NoCrawl,
                ReadOnlyField = field.ReadOnlyField,
                Required = field.Required,
                Scope = field.Scope,
                Title = field.Title,
            };

            var choices = new FieldType[] { FieldType.Choice, FieldType.GridChoice, FieldType.OutcomeChoice };
            if (field.FieldTypeKind == FieldType.DateTime)
            {
                var fieldCast = (FieldDateTime)field;
                fieldCast.EnsureProperties(
                    fc => fc.DisplayFormat,
                    fc => fc.DateTimeCalendarType,
                    fc => fc.FriendlyDisplayFormat);

                fieldModel.DateFieldFormat = fieldCast.DisplayFormat;
            }
            else if (field.FieldTypeKind == FieldType.Text)
            {
                var fieldCast = (FieldText)field;
                fieldCast.EnsureProperties(
                    fc => fc.MaxLength);

                fieldModel.MaxLength = fieldCast.MaxLength;
            }
            else if (field.FieldTypeKind == FieldType.Note)
            {
                var fieldCast = (FieldMultiLineText)field;
                fieldCast.EnsureProperties(
                    fc => fc.AllowHyperlink,
                    fc => fc.NumberOfLines,
                    fc => fc.AppendOnly,
                    fc => fc.RestrictedMode,
                    fc => fc.RichText);

                fieldModel.NumLines = fieldCast.NumberOfLines;
                fieldModel.AppendOnly = fieldCast.AppendOnly;
                fieldModel.RestrictedMode = fieldCast.RestrictedMode;
                fieldModel.RichTextField = fieldCast.RichText;

            }
            else if (field.FieldTypeKind == FieldType.User)
            {
                var fieldCast = (FieldUser)field;
                fieldCast.EnsureProperties(
                    fc => fc.SelectionMode,
                    fc => fc.SelectionGroup,
                    fc => fc.AllowDisplay,
                    fc => fc.Presence,
                    fc => fc.AllowMultipleValues,
                    fc => fc.IsDependentLookup,
                    fc => fc.IsRelationship,
                    fc => fc.LookupList,
                    fc => fc.LookupField,
                    fc => fc.DependentLookupInternalNames,
                    fc => fc.PrimaryFieldId);

                fieldModel.PeopleLookupField = fieldCast.LookupField;
                fieldModel.PeopleOnly = (fieldCast.SelectionMode == FieldUserSelectionMode.PeopleOnly);
                fieldModel.MultiChoice = fieldCast.AllowMultipleValues;

                if (fieldCast.SelectionGroup > 0)
                {
                    var groupId = fieldCast.SelectionGroup;
                    if (siteGroups == null)
                    {
                        siteGroups = fieldWeb.Context.LoadQuery(fieldWeb.SiteGroups
                            .Where(w => w.Id == groupId).Include(group => group.Id, group => group.Title));
                        fieldWeb.Context.ExecuteQueryRetry();
                    }

                    if (siteGroups.Any(sg => sg.Id == groupId))
                    {
                        // we loaded this into context earlier
                        var groupObject = siteGroups.FirstOrDefault(fg => fg.Id == groupId);
                        fieldModel.PeopleGroupName = groupObject.Title;
                    }
                }
            }
            else if (field.FieldTypeKind == FieldType.Lookup)
            {
                var fieldCast = (FieldLookup)field;
                fieldCast.EnsureProperties(
                    fc => fc.LookupList,
                    fc => fc.LookupField,
                    fc => fc.LookupWebId,
                    fc => fc.FromBaseType,
                    fc => fc.AllowMultipleValues,
                    fc => fc.IsDependentLookup,
                    fc => fc.IsRelationship,
                    fc => fc.DependentLookupInternalNames,
                    fc => fc.PrimaryFieldId);

                if (!string.IsNullOrEmpty(fieldCast.LookupList))
                {
                    if (schemaXml == null)
                    {
                        var xdoc = XDocument.Parse(field.SchemaXml, LoadOptions.PreserveWhitespace);
                        schemaXml = xdoc.Element("Field");
                    }

                    var lookupGuid = fieldCast.LookupList.TryParseGuid(Guid.Empty);
                    if (lookupGuid != Guid.Empty)
                    {
                        var fieldLookup = fieldWeb.Lists.GetById(lookupGuid);
                        field.Context.Load(fieldLookup, flp => flp.Title);
                        field.Context.ExecuteQueryRetry();

                        fieldModel.LookupListName = fieldLookup.Title;
                        fieldModel.LookupListFieldName = fieldCast.LookupField;
                    }
                }
                fieldModel.MultiChoice = fieldCast.AllowMultipleValues;
            }
            else if (field.FieldTypeKind == FieldType.Calculated)
            {
                var fieldCast = (FieldCalculated)field;
                fieldCast.EnsureProperties(
                    fc => fc.DateFormat,
                    fc => fc.DisplayFormat,
                    fc => fc.Formula,
                    fc => fc.OutputType,
                    fc => fc.ShowAsPercentage);

                if (schemaXml == null)
                {
                    var xdoc = XDocument.Parse(field.SchemaXml, LoadOptions.PreserveWhitespace);
                    schemaXml = xdoc.Element("Field");
                }

                var xfieldReferences = schemaXml.Element("FieldRefs");
                if (xfieldReferences != null)
                {
                    var fieldreferences = new List<string>();
                    var xFields = xfieldReferences.Elements("FieldRef");
                    if (xFields != null)
                    {

                        foreach (var xField in xFields)
                        {
                            var xFieldName = xField.Attribute("Name");
                            fieldreferences.Add(xFieldName.Value);
                        }
                    }

                    fieldModel.FieldReferences = fieldreferences;
                }

                fieldModel.OutputType = fieldCast.OutputType;
                fieldModel.DefaultFormula = fieldCast.Formula;
            }
            else if (field.FieldTypeKind == FieldType.URL)
            {
                var fieldCast = (FieldUrl)field;
                fieldCast.EnsureProperties(
                    fc => fc.DisplayFormat);

                fieldModel.UrlFieldFormat = fieldCast.DisplayFormat;
            }
            else if (field.FieldTypeKind == FieldType.MultiChoice)
            {
                var fieldCast = (FieldMultiChoice)field;
                fieldCast.EnsureProperties(
                    fc => fc.Choices,
                    fc => fc.DefaultValue,
                    fc => fc.FillInChoice,
                    fc => fc.Mappings);

                var choiceOptions = fieldCast.Choices.Select(s =>
                {
                    var optionDefault = default(Nullable<bool>);
                    if (!string.IsNullOrEmpty(field.DefaultValue)
                        && (field.DefaultValue.Equals(s, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        optionDefault = true;
                    }

                    var option = new SPChoiceModel()
                    {
                        Choice = s,
                        DefaultChoice = optionDefault
                    };
                    return option;
                }).ToList();

                fieldModel.FieldChoices = choiceOptions;
                fieldModel.MultiChoice = field.FieldTypeKind == FieldType.MultiChoice;
            }
            else if (choices.Any(a => a == field.FieldTypeKind))
            {
                var fieldCast = (FieldChoice)field;
                fieldCast.EnsureProperties(
                    fc => fc.Choices,
                    fc => fc.DefaultValue,
                    fc => fc.FillInChoice,
                    fc => fc.Mappings,
                    fc => fc.EditFormat);

                var choiceOptions = fieldCast.Choices.Select(s =>
                {
                    var optionDefault = default(Nullable<bool>);
                    if (!string.IsNullOrEmpty(field.DefaultValue)
                        && (field.DefaultValue.Equals(s, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        optionDefault = true;
                    }

                    var option = new SPChoiceModel()
                    {
                        Choice = s,
                        DefaultChoice = optionDefault
                    };
                    return option;
                }).ToList();

                fieldModel.FieldChoices = choiceOptions;
                fieldModel.ChoiceFormat = fieldCast.EditFormat;
                fieldModel.MultiChoice = field.FieldTypeKind == FieldType.MultiChoice;
            }


            return fieldModel;
        }

    }
}
