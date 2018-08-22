using System.Management.Automation;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WorkflowServices;
using InfrastructureAsCode.Powershell.Commands;
using InfrastructureAsCode.Powershell.Commands.Base;

namespace InfrastructureAsCode.Powershell.Commands.Workflow
{
    [Cmdlet(VerbsCommon.Get, "IaCWorkflowDefinition")]
    public class GetIaCWorkflowDefinition : IaCCmdlet
    {
        [Parameter(Mandatory = false, HelpMessage = "The name of the workflow", Position = 0)]
        public string Name { get; set; }

        [Parameter(Mandatory = false)]
        public SwitchParameter PublishedOnly { get; set; }

        public override void ExecuteCmdlet()
        {
            var SelectedWeb = this.ClientContext.Web;

            if (string.IsNullOrEmpty(Name))
            {
                var servicesManager = new WorkflowServicesManager(ClientContext, SelectedWeb);
                var deploymentService = servicesManager.GetWorkflowDeploymentService();
                var definitions = deploymentService.EnumerateDefinitions(PublishedOnly);

                ClientContext.Load(definitions);

                ClientContext.ExecuteQueryRetry();
                WriteObject(definitions, true);
            }
            else
            {
                WriteObject(SelectedWeb.GetWorkflowDefinition(Name, PublishedOnly));
            }
        }
    }

}
