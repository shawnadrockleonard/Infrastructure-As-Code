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
    using InfrastructureAsCode.Powershell.CmdLets;
    using InfrastructureAsCode.Core.Models;
    using InfrastructureAsCode.Powershell;
    using OfficeDevPnP.Core.Utilities;

    [Cmdlet(VerbsCommon.Get, "IaCListItemCount")]
    [CmdletHelp("Returns the library item count", Category = "ListItems")]
    public class GetIaCListItemCount : IaCCmdlet
    {
        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
        public ListPipeBind Identity { get; set; }


        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();
            try
            {
                var l = Identity.GetList(this.ClientContext.Web);
                l.EnsureProperties(lctx => lctx.ItemCount);
                var itemCount = l.ItemCount;
                LogVerbose(string.Format("The library {0} has {1} items", l.Title, itemCount));
            }
            catch (Exception ex)
            {
                LogError(ex, "Failed in IaCListItemCount for Library {0}", ex.Message);
            }
        }

    }
}
