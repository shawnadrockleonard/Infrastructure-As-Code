using System;
using System.Linq;
using System.Management.Automation;
using Microsoft.SharePoint.Client;
using InfrastructureAsCode.Powershell.PipeBinds;
using InfrastructureAsCode.Powershell.CmdLets;

namespace InfrastructureAsCode.Powershell.Commands.Workflows
{
    [Cmdlet(VerbsLifecycle.Stop, "IaCWorkflowInstance")]
    public class StopIaCWorkflowInstance : IaCCmdlet
    {
        [Parameter(Mandatory = true, HelpMessage = "The instance to stop", Position = 0)]
        public WorkflowInstancePipeBind Identity { get; set; }

        [Parameter(Mandatory = false, HelpMessage = "A list to search the instances for", Position = 1)]
        public ListPipeBind List { get; set; }

        [Parameter(Mandatory = false)]
        public int ListItemId { get; set; }


        public override void ExecuteCmdlet()
        {

            var SelectedWeb = this.ClientContext.Web;


            var list = List.GetList(SelectedWeb);
            var item = list.GetItemById("" + ListItemId);
            list.Context.Load(item, ictx => ictx.Id, ictx => ictx.ParentList.Id);
            list.Context.ExecuteQueryRetry();



            var allinstances = SelectedWeb.GetWorkflowInstances(item);
            foreach (var instance in allinstances.Where(instance => instance.Id == Identity.Id))
            {
                instance.CancelWorkFlow();
                break;
            }
        }
    }


}
