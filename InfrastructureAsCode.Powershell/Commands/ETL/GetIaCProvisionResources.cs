using InfrastructureAsCode.Core.Models;
using InfrastructureAsCode.Powershell.CmdLets;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using OfficeDevPnP.Core.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace InfrastructureAsCode.Powershell.Commands.ETL
{
    /// <summary>
    /// The function cmdlet will query the site specified in the connection and build a configuration file
    /// </summary>
    [Cmdlet(VerbsCommon.Get, "IaCProvisionResources")]
    [CmdletHelp("Get site definition components and write to JSON file.", Category = "ETL")]
    public class GetIaCProvisionResources : IaCCmdlet
    {
        /// <summary>
        /// Represents the directory path for any JSON files for serialization
        /// </summary>
        [Parameter(Mandatory = true, HelpMessage = "Provide a full path to the provisioner JSON file", Position = 0, ValueFromPipeline = true)]
        public string ProvisionerFilePath { get; set; }

        /// <summary>
        /// Specific list to be updated from the above action list
        /// </summary>
        [Parameter(Mandatory = false, ParameterSetName = "ActionDependency")]
        public string[] SpecificLists { get; set; }

        private bool _filterLists { get; set; }

        /// <summary>
        /// Holds the SharePoint groups in the site or created in the site
        /// </summary>
        private List<SPGroupDefinitionModel> siteGroups { get; set; }

        /// <summary>
        /// Holds the [Site] columns
        /// </summary>
        private List<SPFieldDefinitionModel> siteColumns { get; set; }

        /// <summary>
        /// Holds the [List] columns
        /// </summary>
        private List<SPFieldDefinitionModel> listColumns { get; set; }


        /// <summary>
        /// Validate parameters
        /// </summary>
        protected override void OnBeginInitialize()
        {
            var fileinfo = new System.IO.FileInfo(ProvisionerFilePath);

            if (!fileinfo.Directory.Exists)
            {
                throw new System.IO.DirectoryNotFoundException(string.Format("The provisioner directory was not found {0}", fileinfo.DirectoryName));
            }

            _filterLists = (SpecificLists != null && SpecificLists.Any());
            siteGroups = new List<SPGroupDefinitionModel>();
            siteColumns = new List<SPFieldDefinitionModel>();
            listColumns = new List<SPFieldDefinitionModel>();
        }

        /// <summary>
        /// Process the request
        /// </summary>
        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();

            // File Info
            var fileInfo = new System.IO.FileInfo(this.ProvisionerFilePath);

            // SharePoint URI for XML parsing
            XNamespace ns = "http://schemas.microsoft.com/sharepoint/";

            // Skip these specific fields
            var skiptypes = new FieldType[]
            {
                FieldType.Computed,
                FieldType.ContentTypeId,
                FieldType.Invalid,
                FieldType.WorkflowStatus,
                FieldType.WorkflowEventType,
                FieldType.Threading,
                FieldType.ThreadIndex,
                FieldType.Recurrence,
                FieldType.PageSeparator,
                FieldType.OutcomeChoice
            };

            var skipcolumns = new string[]
            {
                "_Hidden",
                "Base Columns",
                "Content Feedback",
                "Core Contact and Calendar Columns",
                "Core Document Columns",
                "Core Task and Issue Columns",
                "Display Template Columns",
                "Document and Record Management Columns",
                "Enterprise Keywords Group",
                "Extended Columns",
                "JavaScript Display Template Columns",
                "Page Layout Columns",
                "Publishing Columns",
                "Reports",
                "Status Indicators",
                "Translation Columns"
            };

            var skipcontenttypes = new string[]
            {
                "_Hidden",
                "Business Intelligence",
                "Community Content Types",
                "Digital Asset Content Types",
                "Display Template Content Types",
                "Document Content Types",
                "Document Set Content Types",
                "Folder Content Types",
                "Content Feedback",
                "Publishing Content Types",
                "Page Layout Content Types",
                "Special Content Types",
                "Group Work Content Types",
                "List Content Types"
            };

            // Construct the model
            var SiteComponents = new SiteProvisionerModel();

            // Load the Context
            var contextWeb = this.ClientContext.Web;
            var fields = this.ClientContext.Web.Fields;
            this.ClientContext.Load(contextWeb, ctxw => ctxw.ServerRelativeUrl, ctxw => ctxw.Id);
            this.ClientContext.Load(fields);

            var groupQuery = this.ClientContext.LoadQuery(contextWeb.SiteGroups
                .Include(group => group.Id,
                        group => group.Title,
                        group => group.Description,
                        group => group.AllowRequestToJoinLeave,
                        group => group.AllowMembersEditMembership,
                        group => group.AutoAcceptRequestToJoinLeave,
                        group => group.OnlyAllowMembersViewMembership,
                        group => group.RequestToJoinLeaveEmailSetting));

            var contentTypes = this.ClientContext.LoadQuery(contextWeb.ContentTypes
                .Include(
                        ict => ict.Id,
                        ict => ict.Group,
                        ict => ict.Description,
                        ict => ict.Name,
                        ict => ict.Hidden,
                        ict => ict.JSLink,
                        ict => ict.FieldLinks,
                        ict => ict.Fields));


            var collists = contextWeb.Lists;
            var lists = this.ClientContext.LoadQuery(collists
                .Include(
                    linc => linc.Title,
                    linc => linc.Id,
                    linc => linc.Description,
                    linc => linc.Hidden,
                    linc => linc.OnQuickLaunch,
                    linc => linc.BaseTemplate,
                    linc => linc.ContentTypesEnabled,
                    linc => linc.AllowContentTypes,
                    linc => linc.EnableFolderCreation,
                    linc => linc.IsApplicationList,
                    linc => linc.IsCatalog,
                    linc => linc.IsSiteAssetsLibrary,
                    linc => linc.IsPrivate,
                    linc => linc.IsSystemList,
                    linc => linc.Views,
                    linc => linc.Fields
                    .Include(
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
                        lft => lft.Hidden,
                        lft => lft.Indexed,
                        lft => lft.InternalName,
                        lft => lft.JSLink,
                        lft => lft.NoCrawl,
                        lft => lft.ReadOnlyField,
                        lft => lft.Required,
                        lft => lft.SchemaXml,
                        lft => lft.Scope,
                        lft => lft.Title
                        ),
                    linc => linc.ContentTypes
                    .Include(
                        ict => ict.Id,
                        ict => ict.Group,
                        ict => ict.Description,
                        ict => ict.Name,
                        ict => ict.Hidden,
                        ict => ict.JSLink,
                        ict => ict.FieldLinks,
                        ict => ict.Fields)).Where(w => !w.IsSystemList && !w.IsSiteAssetsLibrary));
            this.ClientContext.ExecuteQueryRetry();


            if (groupQuery.Any())
            {
                SiteComponents.Groups = new List<SPGroupDefinitionModel>();

                foreach (var group in groupQuery)
                {
                    var model = new SPGroupDefinitionModel()
                    {
                        Id = group.Id,
                        Title = group.Title,
                        Description = group.Description,
                        AllowRequestToJoinLeave = group.AllowRequestToJoinLeave,
                        AllowMembersEditMembership = group.AllowMembersEditMembership,
                        AutoAcceptRequestToJoinLeave = group.AutoAcceptRequestToJoinLeave,
                        OnlyAllowMembersViewMembership = group.OnlyAllowMembersViewMembership,
                        RequestToJoinLeaveEmailSetting = group.RequestToJoinLeaveEmailSetting
                    };

                    SiteComponents.Groups.Add(model);
                }
            }



            if (fields.Any())
            {
                var webfields = new List<SPFieldDefinitionModel>();
                foreach (Microsoft.SharePoint.Client.Field field in fields)
                {
                    if (skiptypes.Any(st => field.FieldTypeKind == st)
                        || skipcolumns.Any(sg => field.Group.Equals(sg, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        continue;
                    }

                    try
                    {
                        var fieldModel = RetrieveField(field, groupQuery);
                        webfields.Add(fieldModel);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Trace.TraceError("Failed to parse field {0} MSG:{1}", field.InternalName, ex.Message);
                    }
                }

                SiteComponents.FieldDefinitions = webfields;
            }


            var contentTypesFieldset = new List<dynamic>();
            if (contentTypes.Any())
            {
                SiteComponents.ContentTypes = new List<SPContentTypeDefinition>();
                foreach (ContentType contenttype in contentTypes)
                {
                    // skip core content types
                    if (skipcontenttypes.Any(sg => contenttype.Group.Equals(sg, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        continue;
                    }

                    contenttype.EnsureProperties(ctp => ctp.Id, ctp => ctp.Group, ctp => ctp.Hidden, ctp => ctp.Description, ctp => ctp.Name, ctp => ctp.FieldLinks);

                    var ctypemodel = new SPContentTypeDefinition()
                    {
                        ContentTypeId = contenttype.Id.StringValue,
                        ContentTypeGroup = contenttype.Group,
                        Hidden = contenttype.Hidden,
                        Description = contenttype.Description,
                        Name = contenttype.Name
                    };

                    if (contenttype.FieldLinks.Any())
                    {
                        ctypemodel.FieldLinks = new List<SPFieldLinkDefinitionModel>();
                        foreach (FieldLink fieldlink in contenttype.FieldLinks)
                        {
                            ctypemodel.FieldLinks.Add(new SPFieldLinkDefinitionModel()
                            {
                                Id = fieldlink.Id,
                                Name = fieldlink.Name,
                                Required = fieldlink.Required,
                                Hidden = fieldlink.Hidden
                            });

                            contentTypesFieldset.Add(new { ctypeid = contenttype.Id.StringValue, name = fieldlink.Name });
                        }
                    }

                    SiteComponents.ContentTypes.Add(ctypemodel);
                }
            }


            if (lists.Any())
            {
                SiteComponents.Lists = new List<SPListDefinition>();

                foreach (List list in lists.Where(lwt =>
                    (!_filterLists
                        || (_filterLists && SpecificLists.Any(sl => lwt.Title.Equals(sl, StringComparison.InvariantCultureIgnoreCase))))))
                {
                    var listdefinition = new SPListDefinition()
                    {
                        Id = list.Id,
                        ListName = list.Title,
                        ListDescription = list.Description,
                        QuickLaunch = list.OnQuickLaunch ? QuickLaunchOptions.On : QuickLaunchOptions.Off,
                        ContentTypeEnabledOverride = list.ContentTypesEnabled,
                        EnableFolderCreation = list.EnableFolderCreation,
                        Hidden = list.Hidden,
                        IsApplicationList = list.IsApplicationList,
                        IsCatalog = list.IsCatalog,
                        IsSiteAssetsLibrary = list.IsSiteAssetsLibrary,
                        IsPrivate = list.IsPrivate,
                        IsSystemList = list.IsSystemList
                    };


                    if (list.ContentTypes.Any())
                    {
                        listdefinition.ContentTypes = new List<SPContentTypeDefinition>();
                        foreach (var contenttype in list.ContentTypes)
                        {
                            var ctypemodel = new SPContentTypeDefinition()
                            {
                                Inherits = true,
                                ContentTypeId = contenttype.Id.StringValue,
                                ContentTypeGroup = contenttype.Group,
                                Description = contenttype.Description,
                                Name = contenttype.Name,
                                Hidden = contenttype.Hidden,
                                JSLink = contenttype.JSLink
                            };

                            if (contenttype.FieldLinks.Any())
                            {
                                ctypemodel.FieldLinks = new List<SPFieldLinkDefinitionModel>();
                                foreach (var cfield in contenttype.FieldLinks)
                                {
                                    ctypemodel.FieldLinks.Add(new SPFieldLinkDefinitionModel()
                                    {
                                        Id = cfield.Id,
                                        Name = cfield.Name,
                                        Hidden = cfield.Hidden,
                                        Required = cfield.Required
                                    });

                                }
                            }

                            if (contenttype.Fields.Any())
                            {
                                foreach (var cfield in contenttype.Fields.Where(cf => !ctypemodel.FieldLinks.Any(fl => fl.Name == cf.InternalName)))
                                {
                                    ctypemodel.FieldLinks.Add(new SPFieldLinkDefinitionModel()
                                    {
                                        Id = cfield.Id,
                                        Name = cfield.InternalName,
                                        Hidden = cfield.Hidden,
                                        Required = cfield.Required
                                    });
                                }
                            }

                            listdefinition.ContentTypes.Add(ctypemodel);
                        }
                    }

                    if (list.Fields.Any())
                    {
                        var listfields = new List<SPFieldDefinitionModel>();
                        foreach (Field listField in list.Fields)
                        {
                            // skip internal fields
                            if (skiptypes.Any(st => listField.FieldTypeKind == st)
                                || skipcolumns.Any(sg => listField.Group.Equals(sg, StringComparison.CurrentCultureIgnoreCase)))
                            {
                                continue;
                            }

                            // skip fields that are defined
                            if (contentTypesFieldset.Any(ft => ft.name == listField.InternalName))
                            {
                                continue;
                            }

                            var fieldXml = listField.SchemaXml;
                            if (!string.IsNullOrEmpty(fieldXml))
                            {
                                var xdoc = XDocument.Parse(fieldXml, LoadOptions.PreserveWhitespace);
                                var xField = xdoc.Element("Field");
                                var xSourceID = xField.Attribute("SourceID");
                                var xScope = xField.Element("Scope");
                                if (xSourceID != null && xSourceID.Value.IndexOf(ns.NamespaceName, StringComparison.CurrentCultureIgnoreCase) < 0)
                                {
                                    try
                                    {
                                        var customField = RetrieveField(listField, groupQuery, xField);
                                        listfields.Add(customField);
                                    }
                                    catch (Exception ex)
                                    {
                                        System.Diagnostics.Trace.TraceError("Failed to parse field {0} MSG:{1}", listField.InternalName, ex.Message);
                                    }
                                }
                            }
                        }

                        listdefinition.FieldDefinitions = listfields;
                    }


                    SiteComponents.Lists.Add(listdefinition);
                }
            }

            // Write the JSON to disc
            var jsonsettings = new JsonSerializerSettings()
            {
                Formatting = Formatting.Indented,
                Culture = System.Globalization.CultureInfo.CurrentUICulture,
                DateFormatHandling = DateFormatHandling.IsoDateFormat,
                NullValueHandling = NullValueHandling.Ignore
            };

            var json = JsonConvert.SerializeObject(SiteComponents, jsonsettings);
            System.IO.File.WriteAllText(fileInfo.FullName, json);
        }

        /// <summary>
        /// Parse the Field into the appropriate object
        /// </summary>
        /// <param name="field"></param>
        /// <param name="siteGroups"></param>
        /// <param name="schemaXml">(OPTIONAL) the schema xml parsed into an XDocument</param>
        /// <returns></returns>
        private SPFieldDefinitionModel RetrieveField(Microsoft.SharePoint.Client.Field field, IEnumerable<Group> siteGroups = null, XElement schemaXml = null)
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

                if (fieldCast.SelectionGroup > 0
                    && (siteGroups != null && siteGroups.Any(sg => sg.Id == fieldCast.SelectionGroup)))
                {
                    // we loaded this into context earlier
                    var groupObject = siteGroups.FirstOrDefault(fg => fg.Id == fieldCast.SelectionGroup);
                    fieldModel.PeopleGroupName = groupObject.Title;
                }
            }
            else if (field.FieldTypeKind == FieldType.Lookup)
            {
                var fieldCast = (FieldLookup)field;
                fieldCast.EnsureProperties(
                    fc => fc.LookupList,
                    fc => fc.LookupField,
                    fc => fc.AllowMultipleValues,
                    fc => fc.IsDependentLookup,
                    fc => fc.IsRelationship,
                    fc => fc.DependentLookupInternalNames,
                    fc => fc.PrimaryFieldId);

                if (!string.IsNullOrEmpty(fieldCast.LookupList))
                {
                    var lookupGuid = new Guid(fieldCast.LookupList);
                    var fieldLookup = this.ClientContext.Web.Lists.GetById(lookupGuid);
                    this.ClientContext.Load(fieldLookup, flp => flp.Title);
                    this.ClientContext.ExecuteQueryRetry();

                    fieldModel.LookupListName = fieldLookup.Title;
                    fieldModel.LookupListFieldName = fieldCast.LookupField;
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
