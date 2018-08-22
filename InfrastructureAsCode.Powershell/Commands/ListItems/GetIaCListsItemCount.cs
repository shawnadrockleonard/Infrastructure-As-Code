using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Powershell.Commands.ListItems
{
    using Microsoft.SharePoint.Client;
    using InfrastructureAsCode.Powershell.PipeBinds;
    using InfrastructureAsCode.Powershell.Commands.Base;
    using InfrastructureAsCode.Core.Models;
    using InfrastructureAsCode.Powershell;


    [Cmdlet(VerbsCommon.Get, "IaCListsItemCount")]
    [CmdletHelp("Returns all lists for the web and its related item count", Category = "ListItems")]
    public class GetIaCListsItemCount : IaCCmdlet
    {
        [Parameter(Mandatory = false, ValueFromPipeline = true, Position = 0)]
        public ListPipeBind Identity;

        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();
            try
            {

                var listCollection = ClientContext.Web.Lists.Include(lctx => lctx.ItemCount, lctx => lctx.Id, lctx => lctx.Title);
                ClientContext.LoadQuery(listCollection);
                ClientContext.ExecuteQueryRetry();

                foreach (var lItem in listCollection)
                {
                    LogVerbose("This list {0} has this many items {1}", lItem.Title, lItem.ItemCount);
                }
            }
            catch (Exception ex)
            {
                LogError(ex, "Failed in IaCListsItemCount for Library {0}", ex.Message);
            }
        }

    }
}
