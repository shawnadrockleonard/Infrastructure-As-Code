using System.Management.Automation;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WorkflowServices;
using InfrastructureAsCode.Powershell.PipeBinds;
using InfrastructureAsCode.Powershell.CmdLets;

namespace InfrastructureAsCode.Powershell.Commands.Workflows
{
    [Cmdlet(VerbsCommon.Get, "IaCWorkflowSubscription")]
    public class GetIaCWorkflowSubscription : IaCCmdlet
    {
        [Parameter(Mandatory = false, HelpMessage = "The name of the workflow", Position = 0)]
        public string Name;

        [Parameter(Mandatory = false, HelpMessage = "A list to search the association for", Position = 1)]
        public ListPipeBind List;

        public override void ExecuteCmdlet()
        {
            var SelectedWeb = this.ClientContext.Web;

            if (List != null)
            {
                var list = List.GetList(SelectedWeb);

                if (string.IsNullOrEmpty(Name))
                {
                    var servicesManager = new WorkflowServicesManager(ClientContext, SelectedWeb);
                    var subscriptionService = servicesManager.GetWorkflowSubscriptionService();
                    WorkflowSubscriptionCollection subscriptions = subscriptionService.EnumerateSubscriptionsByList(list.Id);

                    ClientContext.Load(subscriptions);
                    ClientContext.ExecuteQueryRetry();

                    WriteObject(subscriptions, true);
                }
                else
                {
                    WriteObject(list.GetWorkflowSubscription(Name));
                }
            }
            else
            {
                if (string.IsNullOrEmpty(Name))
                {
                    var servicesManager = new WorkflowServicesManager(ClientContext, SelectedWeb);
                    var subscriptionService = servicesManager.GetWorkflowSubscriptionService();
                    var subscriptions = subscriptionService.EnumerateSubscriptions();

                    ClientContext.Load(subscriptions);
                    ClientContext.ExecuteQueryRetry();

                    WriteObject(subscriptions, true);
                }
                else
                {
                    WriteObject(SelectedWeb.GetWorkflowSubscription(Name));
                }
            }
        }
    }

}
