using IaC.Powershell.CmdLets;
using IaC.Powershell.Extensions;
using IaC.Powershell.PipeBinds;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;

namespace IaC.Powershell.Commands.Lists
{
    [Cmdlet(VerbsCommon.Get, "IaCListItemCount")]
    [CmdletHelp("Returns the library item count", Category = "ListItems")]
    public class GetIaCListItemCount : IaCCmdlet
    {
        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
        public string LibraryName;

        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();
            try
            {
                var ctx = this.ClientContext;

                var w = ctx.Web;
                var l = w.Lists.GetByTitle(LibraryName);
                ctx.Load(w);
                ctx.Load(l);
                ClientContext.ExecuteQueryRetry();

                var itemCount = l.ItemCount;
                LogVerbose(string.Format("The library {0} has {1} items", LibraryName, itemCount));

                var listCollection = w.Lists;
                ctx.Load(listCollection, ll => ll.Include(p => p.Title, p => p.Id, pp => pp.ItemCount));
                ClientContext.ExecuteQueryRetry();

                foreach (var lItem in listCollection)
                {
                    LogVerbose("This list {0} has this many items {1}", lItem.Title, lItem.ItemCount);
                }
            }
            catch (Exception ex)
            {
                LogError(ex, "Failed in GetListItemCount for Library {0}", LibraryName);
            }
        }

    }
}
