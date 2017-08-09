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
    /// The function cmdlet will upgrade the EzForms site specified in the connection to the latest configuration changes
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
        public string SpecificListName { get; set; }

        /// <summary>
        /// Specific view to be updated from the above action list
        /// </summary>
        [Parameter(Mandatory = false, ParameterSetName = "ActionDependency")]
        public string SpecificViewName { get; set; }


        internal List<SPGroupDefinitionModel> siteGroups { get; set; }

        internal List<SPFieldDefinitionModel> siteColumns { get; set; }


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

            siteGroups = new List<SPGroupDefinitionModel>();
            siteColumns = new List<SPFieldDefinitionModel>();
        }

        /// <summary>
        /// Process the request
        /// </summary>
        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();

            // File Info
            var fileInfo = new System.IO.FileInfo(this.ProvisionerFilePath);

            // Skip these specific fields
            var skiptypes = new FieldType[] {
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

            // Construct the model
            var SiteComponents = new SiteProvisionerModel();

            // Site Columns
            var contextWeb = this.ClientContext.Web;
            var fields = this.ClientContext.Web.Fields;
            var groupQuery = this.ClientContext.Web.SiteGroups;
            var contentTypes = this.ClientContext.Web.ContentTypes;
            this.ClientContext.Load(contextWeb, ctxw => ctxw.ServerRelativeUrl, ctxw => ctxw.Id);
            this.ClientContext.Load(fields);
            this.ClientContext.Load(groupQuery);
            this.ClientContext.Load(contentTypes);
            this.ClientContext.ExecuteQueryRetry();


            if (fields.Any())
            {
                var webfields = new List<SPFieldDefinitionModel>();
                foreach (Microsoft.SharePoint.Client.Field field in fields)
                {
                    if (skiptypes.Any(st => field.FieldTypeKind == st))
                    {
                        continue;
                    }

                    var fieldModel = RetrieveField(field);
                    webfields.Add(fieldModel);
                }

                SiteComponents.FieldDefinitions = webfields;
            }

            if (contentTypes.Any())
            {
                SiteComponents.ContentTypes = new List<SPContentTypeDefinition>();
                foreach (ContentType contenttype in contentTypes)
                {
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
                        }
                    }

                    SiteComponents.ContentTypes.Add(ctypemodel);
                }
            }


            var collists = contextWeb.Lists;
            var lists = this.ClientContext.LoadQuery(collists.Include(linc => linc.Title,
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
                        linc => linc.Fields,
                        linc => linc.ContentTypes).Where(w => !w.IsSystemList && !w.IsSiteAssetsLibrary));
            this.ClientContext.ExecuteQueryRetry();

            if (lists.Any())
            {
                SiteComponents.Lists = new List<SPListDefinition>();

                foreach (List list in lists.Where(lwt =>
                    (string.IsNullOrEmpty(SpecificListName)
                        || (!String.IsNullOrEmpty(SpecificListName) && lwt.Title.Equals(SpecificListName, StringComparison.InvariantCultureIgnoreCase)))))
                {
                    var listdefinition = new SPListDefinition()
                    {
                        Id = list.Id,
                        ListName = list.Title,
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

                    if(list.ContentTypes.Any())
                    {
                        listdefinition.ContentTypes = new List<SPContentTypeDefinition>();
                        foreach(var contenttype in list.ContentTypes)
                        {
                            listdefinition.ContentTypes.Add(new SPContentTypeDefinition()
                            {
                                Inherits = true,
                                ContentTypeId = contenttype.Id.StringValue,
                                ContentTypeGroup = contenttype.Group,
                                Description = contenttype.Description,
                                Name = contenttype.Name,
                                Hidden = contenttype.Hidden,
                                JSLink = contenttype.JSLink
                            });
                        }
                    }

                    if (list.Fields.Any())
                    {
                        var listfields = new List<SPFieldDefinitionModel>();
                        XNamespace ns = "http://schemas.microsoft.com/sharepoint/";
                        foreach (var listField in list.Fields)
                        {
                            if (skiptypes.Any(st => listField.FieldTypeKind == st))
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
                                    var customField = RetrieveField(listField);
                                    listfields.Add(customField);
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

        private SPFieldDefinitionModel RetrieveField(Microsoft.SharePoint.Client.Field field)
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

            var choices = new FieldType[] { FieldType.Choice, FieldType.GridChoice, FieldType.MultiChoice, FieldType.OutcomeChoice };
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
                if (fieldCast.SelectionMode == FieldUserSelectionMode.PeopleAndGroups && fieldCast.SelectionGroup > 0)
                {
                    var groupObject = this.ClientContext.Web.SiteGroups.GetById(fieldCast.SelectionGroup);
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
                    try
                    {
                        var lookupGuid = new Guid(fieldCast.LookupList);
                        var fieldLookup = this.ClientContext.Web.Lists.GetById(lookupGuid);
                        this.ClientContext.Load(fieldLookup, flp => flp.Title);
                        this.ClientContext.ExecuteQueryRetry();

                        fieldModel.LookupListName = fieldLookup.Title;
                        fieldModel.LookupListFieldName = fieldCast.LookupField;
                    }
                    catch(Exception ex)
                    {
                        System.Diagnostics.Trace.TraceError("Failed to pull lookup list {0} MSG:{1}", fieldCast.LookupList, ex.Message);
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

            }
            else if (field.FieldTypeKind == FieldType.URL)
            {
                var fieldCast = (FieldUrl)field;
                fieldCast.EnsureProperties(
                    fc => fc.DisplayFormat);

                fieldModel.UrlFieldFormat = fieldCast.DisplayFormat;
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
