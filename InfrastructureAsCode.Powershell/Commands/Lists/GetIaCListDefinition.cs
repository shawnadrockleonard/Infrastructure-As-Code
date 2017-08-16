using InfrastructureAsCode.Core.Extensions;
using InfrastructureAsCode.Core.Models;
using InfrastructureAsCode.Core.Models.Minimal;
using InfrastructureAsCode.Powershell.CmdLets;
using InfrastructureAsCode.Powershell.PipeBinds;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Utilities;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Powershell.Commands.Lists
{
    /// <summary>
    /// Returns the list definition, views, columns, settings
    /// </summary>
    /// <remarks>
    /// Get-IaCListDefinition -List ""Demo List""
    /// </remarks>
    [Cmdlet(VerbsCommon.Get, "IaCListDefinition")]
    [OutputType(typeof(SPListDefinition))]
    public class GetIaCListDefinition : IaCCmdlet
    {
        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0, HelpMessage = "The ID or Url of the list.")]
        public ListPipeBind Identity;

        /// <summary>
        /// Expand the list definition
        /// </summary>
        [Parameter(Mandatory = false, Position = 1)]
        public SwitchParameter ExpandObjects { get; set; }

        /// <summary>
        /// Extract the list items
        /// </summary>
        [Parameter(Mandatory = false, Position = 2)]
        public SwitchParameter ExtractData { get; set; }


        public override void ExecuteCmdlet()
        {

            if (Identity != null)
            {
                var list = Identity.GetList(this.ClientContext.Web);
                if (list != null)
                {
                    // ---> Site Usage Properties
                    var _ctx = this.ClientContext;
                    var _site = _ctx.Site;


                    ClientContext.Load(_site, cts => cts.Usage, cts => cts.Id);


                    ClientContext.Load(list,
                        lctx => lctx.Id,
                        lctx => lctx.Title,
                        lctx => lctx.Description,
                        lctx => lctx.DefaultViewUrl,
                        lctx => lctx.Created,
                        lctx => lctx.LastItemModifiedDate,
                        lctx => lctx.LastItemUserModifiedDate,
                        lctx => lctx.EnableModeration,
                        lctx => lctx.EnableVersioning,
                        lctx => lctx.CreatablesInfo,
                        lctx => lctx.EnableVersioning,
                        lctx => lctx.RootFolder.ServerRelativeUrl);
                    ClientContext.ExecuteQueryRetry();



                    var listmodel = new SPListDefinition()
                    {
                        Id = list.Id,
                        ListName = list.Title,
                        ServerRelativeUrl = list.DefaultViewUrl,
                        Created = list.Created,
                        LastItemModifiedDate = list.LastItemModifiedDate,
                        LastItemUserModifiedDate = list.LastItemUserModifiedDate
                    };


                    if (ExpandObjects)
                    {
                        var views = ClientContext.LoadQuery(list.Views
                            .IncludeWithDefaultProperties(v => v.ViewFields));

                        var fields = ClientContext.LoadQuery(list.Fields.IncludeWithDefaultProperties(
                            v => v.AutoIndexed, v => v.CanBeDeleted, v => v.DefaultFormula, v => v.DefaultValue, v => v.Description, v => v.EnforceUniqueValues,
                            v => v.FieldTypeKind, v => v.Filterable, v => v.Group, v => v.Hidden, v => v.Id, v => v.InternalName, v => v.Indexed, v => v.JSLink, v => v.NoCrawl, v => v.ReadOnlyField,
                            v => v.Required, v => v.Title));
                        ClientContext.ExecuteQueryRetry();

                        foreach (var field in fields)
                        {
                            listmodel.FieldDefinitions.Add(new SPFieldDefinitionModel()
                            {
                                FieldGuid = field.Id,
                                AutoIndexed = field.AutoIndexed,
                                CanBeDeleted = field.CanBeDeleted,
                                DefaultFormula = field.DefaultFormula,
                                DefaultValue = field.DefaultValue,
                                Description = field.Description,
                                EnforceUniqueValues = field.EnforceUniqueValues,
                                FieldTypeKind = field.FieldTypeKind,
                                Filterable = field.Filterable,
                                GroupName = field.Group,
                                HiddenField = field.Hidden,
                                InternalName = field.InternalName,
                                FieldIndexed = field.Indexed,
                                JSLink = field.JSLink,
                                NoCrawl = field.NoCrawl,
                                ReadOnlyField = field.ReadOnlyField,
                                Required = field.Required,
                                Title = field.Title
                            });
                        }
                    }

                    if(ExtractData)
                    {
                        listmodel.ListItems = new List<SPListItemDefinition>();

                        ListItemCollectionPosition itemPosition = null;
                        var camlQuery = new CamlQuery()
                        {
                            ViewXml = CAML.ViewQuery(ViewScope.RecursiveAll,
                                        string.Empty,
                                        CAML.OrderBy(new OrderByField("Title")),
                                        CAML.ViewFields((new string[] { "Title", "Author", "Created", "Editor", "Modified" }).Select(s => CAML.FieldRef(s)).ToArray()),
                                        50)
                        };

                        try
                        {
                            while (true)
                            {
                                camlQuery.ListItemCollectionPosition = itemPosition;
                                ListItemCollection listItems = list.GetItems(camlQuery);
                                this.ClientContext.Load(listItems);
                                this.ClientContext.ExecuteQueryRetry();
                                itemPosition = listItems.ListItemCollectionPosition;

                                foreach (var rbiItem in listItems)
                                {

                                    LogVerbose("Title: {0}; Item ID: {1}", rbiItem["Title"], rbiItem.Id);

                                    var author = rbiItem.RetrieveListItemUserValue("Author");
                                    var editor = rbiItem.RetrieveListItemUserValue("Editor");

                                    var newitem = new SPListItemDefinition()
                                    {
                                        Title = rbiItem.RetrieveListItemValue("Title"),
                                        Id = rbiItem.Id,
                                        Created = rbiItem.RetrieveListItemValue("Created").ToNullableDatetime(),
                                        CreatedBy = new SPPrincipalUserDefinition() {
                                            Id = author.LookupId,
                                            LoginName = author.LookupValue,
                                            Email = author.Email
                                        },
                                        Modified = rbiItem.RetrieveListItemValue("Modified").ToNullableDatetime(),
                                        ModifiedBy = new SPPrincipalUserDefinition()
                                        {
                                            Id = editor.LookupId,
                                            LoginName = editor.LookupValue,
                                            Email = editor.Email
                                        }
                                    };

                                    listmodel.ListItems.Add(newitem);
                                }

                                if (itemPosition == null)
                                {
                                    break;
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            LogError(ex, "Failed to retrieve list item collection");
                        }
                    }

                    WriteObject(listmodel);
                }
            }

        }
    }
}
