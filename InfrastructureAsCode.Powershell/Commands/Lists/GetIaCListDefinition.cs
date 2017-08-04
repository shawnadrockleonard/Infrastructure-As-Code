using InfrastructureAsCode.Core.Extensions;
using InfrastructureAsCode.Core.Models;
using InfrastructureAsCode.Core.Models.Minimal;
using InfrastructureAsCode.Powershell.CmdLets;
using InfrastructureAsCode.Powershell.PipeBinds;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
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

        public override void ExecuteCmdlet()
        {

            if (Identity != null)
            {
                var list = Identity.GetList(this.ClientContext.Web);
                if (list != null)
                {
                    var listmodel = new SPListDefinition()
                    {
                        Id = list.Id,
                        ListName = list.Title,
                        ServerRelativeUrl = list.DefaultViewUrl
                    };

                    var views = ClientContext.LoadQuery(list.Views.IncludeWithDefaultProperties(v => v.ViewFields));

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

                    WriteObject(listmodel);
                }
            }

        }
    }
}
