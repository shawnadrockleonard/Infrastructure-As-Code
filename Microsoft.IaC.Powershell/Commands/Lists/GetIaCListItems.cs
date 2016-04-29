using IaC.Core.Extensions;
using IaC.Core.Models;
using IaC.Powershell.CmdLets;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;

namespace IaC.Powershell.Commands.Lists
{
    [Cmdlet(VerbsCommon.Get, "IaCListItems", SupportsShouldProcess = true)]
    [CmdletHelp("Query the list and output the list items.", Category = "ListItems")]
    public class GetIaCListItems : IaCCmdlet
    {
        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
        public string ListTitle { get; set; }

        /// <summary>
        /// Execute the cmdlet
        /// </summary>
        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();

            Collection<SPListItemDefinition> results = new Collection<SPListItemDefinition>();

            var _listSites = this.ClientContext.Web.Lists.GetByTitle(ListTitle);
            this.ClientContext.Load(_listSites);
            this.ClientContext.ExecuteQuery();

            ListItemCollectionPosition itemPosition = null;

            try
            {
                while (true)
                {
                    CamlQuery camlQuery = new CamlQuery();
                    camlQuery.ListItemCollectionPosition = itemPosition;
                    camlQuery.ViewXml = @"<View><Query>
                                            <ViewFields>
                                                <FieldRef Name='Title'/>
                                             </ViewFields>
                                            <RowLimit>50</RowLimit>
                                        </Query></View>";
                    ListItemCollection listItems = _listSites.GetItems(camlQuery);
                    this.ClientContext.Load(listItems);
                    this.ClientContext.ExecuteQuery();
                    itemPosition = listItems.ListItemCollectionPosition;

                    foreach (var rbiItem in listItems)
                    {

                        LogVerbose("Title: {0}; Item ID: {1}", rbiItem["Title"], rbiItem.Id);
                        var newitem = new SPListItemDefinition()
                        {
                            Title = rbiItem.RetrieveListItemValue("Title"),
                            ID = rbiItem.Id
                        };
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

