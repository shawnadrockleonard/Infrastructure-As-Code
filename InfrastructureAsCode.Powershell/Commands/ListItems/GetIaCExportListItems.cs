using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Powershell.Commands.ListItems
{
    using Microsoft.SharePoint.Client;
    using InfrastructureAsCode.Powershell.PipeBinds;
    using InfrastructureAsCode.Powershell.Commands.Base;
    using OfficeDevPnP.Core.Utilities;

    /// <summary>
    /// Query the sharepoint list, returning the list item ID/Title and export to memory stream
    /// </summary>
    [Cmdlet(VerbsCommon.Get, "IaCExportListItems", SupportsShouldProcess = true)]
    public class GetIaCExportListItems : IaCCmdlet
    {
        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
        public ListPipeBind LibraryName { get; set; }


        /// <summary>
        /// Execute the cmdlet
        /// </summary>
        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();

            Collection<PSObject> results = new Collection<PSObject>();

            var _listSites = LibraryName.GetList( this.ClientContext.Web);

            ListItemCollectionPosition itemPosition = null;

            try
            {
                while (true)
                {
                    CamlQuery camlQuery = new CamlQuery
                    {
                        ListItemCollectionPosition = itemPosition,
                        ViewXml = CAML.ViewQuery(ViewScope.RecursiveAll, string.Empty, string.Empty, CAML.ViewFields(CAML.FieldRef("Title")), 50)
                    };

                    ListItemCollection listItems = _listSites.GetItems(camlQuery);
                    this.ClientContext.Load(listItems);
                    this.ClientContext.ExecuteQueryRetry();
                    itemPosition = listItems.ListItemCollectionPosition;

                    foreach (var rbiItem in listItems)
                    {

                        LogVerbose("Title: {0}; Item ID: {1}", rbiItem["Title"], rbiItem.Id);
                        var newitem = new
                        {
                            Title = rbiItem["Title"],
                            Id = rbiItem.Id
                        };
                        var ps = PSObject.AsPSObject(newitem);
                        results.Add(ps);
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

