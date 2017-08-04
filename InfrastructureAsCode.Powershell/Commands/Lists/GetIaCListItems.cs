using InfrastructureAsCode.Core.Extensions;
using InfrastructureAsCode.Core.Models;
using InfrastructureAsCode.Powershell.CmdLets;
using InfrastructureAsCode.Powershell.PipeBinds;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Utilities;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Powershell.Commands.Lists
{
    /// <summary>
    /// Get List Items for the specifed list
    /// </summary>
    /// <remarks>
    /// Get-IaCListItems -Identity /Lists/Announcements
    /// Get-IaCListItems -Identity "Announcements"
    /// </remarks>
    [Cmdlet(VerbsCommon.Get, "IaCListItems", SupportsShouldProcess = true)]
    [CmdletHelp("Query the list and output the list items.", Category = "ListItems")]
    [OutputType(typeof(Collection<SPListItemDefinition>))]
    public class GetIaCListItems : IaCCmdlet
    {
        [Parameter(Mandatory = false, ValueFromPipeline = true, Position = 0, HelpMessage = "The ID or Url of the list.")]
        public ListPipeBind Identity;

        /// <summary>
        /// Should we expand the list item to include its column data
        /// </summary>
        [Parameter(Mandatory = false)]
        public SwitchParameter Expand { get; set; }

        /// <summary>
        /// Execute the cmdlet
        /// </summary>
        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();

            Collection<SPListItemDefinition> results = new Collection<SPListItemDefinition>();


            var _list = Identity.GetList(this.ClientContext.Web);


            ListItemCollectionPosition itemPosition = null;
            var camlQuery = new CamlQuery()
            {
                ViewXml = CAML.ViewQuery(ViewScope.RecursiveAll,
                            string.Empty,
                            CAML.OrderBy(new OrderByField("Title")),
                            CAML.ViewFields((new string[] { "Title", "Author", "Editor" }).Select(s => CAML.FieldRef(s)).ToArray()),
                            50)
            };

            try
            {
                while (true)
                {
  
                    camlQuery.ListItemCollectionPosition = itemPosition;
                    ListItemCollection listItems = _list.GetItems(camlQuery);
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
                            Id = rbiItem.Id
                        };

                        foreach(var rbiField in rbiItem.FieldValues)
                        {
                            newitem.ColumnValues.Add(new SPListItemFieldDefinition()
                            {
                                FieldName = rbiField.Key,
                                FieldValue = rbiField.Value
                            });
                        }

                        results.Add(newitem);
                    }

                    if (itemPosition == null)
                    {
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                LogError(ex, "Failed to retrieve recycle bin collection");
            }

            WriteObject(results, true);
        }
    }
}

