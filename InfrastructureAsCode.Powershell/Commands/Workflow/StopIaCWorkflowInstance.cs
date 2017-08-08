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
        public WorkflowInstancePipeBind Identity;

        public override void ExecuteCmdlet()
        {
            if (Identity.Instance != null)
            {
                Identity.Instance.CancelWorkFlow();
            }
            else if (Identity.Id != Guid.Empty)
            {
                var SelectedWeb = this.ClientContext.Web;

                var allinstances = SelectedWeb.GetWorkflowInstances();
                foreach (var instance in allinstances.Where(instance => instance.Id == Identity.Id))
                {
                    instance.CancelWorkFlow();
                    break;
                }
            }
        }
    }


}
