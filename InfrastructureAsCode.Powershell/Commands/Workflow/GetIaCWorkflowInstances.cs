using System.Management.Automation;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WorkflowServices;
using InfrastructureAsCode.Powershell.PipeBinds;
using InfrastructureAsCode.Powershell.CmdLets;

namespace InfrastructureAsCode.Powershell.Commands.Workflows
{
    /// <summary>
    /// https://msdn.microsoft.com/en-us/library/office/dn481315.aspx
    /// </summary>
    [Cmdlet(VerbsCommon.Get, "IaCWorkflowInstances")]
    public class GetIaCWorkflowInstances : IaCCmdlet
    {
        [Parameter(Mandatory = false, HelpMessage = "The name of the workflow", Position = 0)]
        public string Name;

        [Parameter(Mandatory = false, HelpMessage = "A list to search the instances for", Position = 1)]
        public ListPipeBind List;

        [Parameter(Mandatory = false)]
        public SwitchParameter PublishedOnly = true;

        public override void ExecuteCmdlet()
        {
            var SelectedWeb = this.ClientContext.Web;

            if (List != null)
            {
                var list = List.GetList(SelectedWeb);

                var servicesManager = new WorkflowServicesManager(ClientContext, SelectedWeb);

                var subscriptionService = servicesManager.GetWorkflowSubscriptionService();
                var subscriptions = subscriptionService.EnumerateSubscriptionsByList(list.Id);

                ClientContext.Load(subscriptions);
                ClientContext.ExecuteQueryRetry();

                using (var subenumerator = subscriptions.GetEnumerator())
                {
                    var deploymentService = servicesManager.GetWorkflowInstanceService();

                    while (subenumerator.MoveNext())
                    {
                        var subscription = subenumerator.Current;
                        if (subscription.Name.Equals(Name, System.StringComparison.CurrentCultureIgnoreCase))
                        {

                            var countTerminated = deploymentService.CountInstancesWithStatus(subscription, WorkflowStatus.Terminated);
                            var countSuspended = deploymentService.CountInstancesWithStatus(subscription, WorkflowStatus.Suspended);
                            var countInvalid = deploymentService.CountInstancesWithStatus(subscription, WorkflowStatus.Invalid);
                            var countCancelled = deploymentService.CountInstancesWithStatus(subscription, WorkflowStatus.Canceled);
                            var countCanceling = deploymentService.CountInstancesWithStatus(subscription, WorkflowStatus.Canceling);
                            var countStarted = deploymentService.CountInstancesWithStatus(subscription, WorkflowStatus.Started);
                            var countNotStarted = deploymentService.CountInstancesWithStatus(subscription, WorkflowStatus.NotStarted);
                            var countNotSpecified = deploymentService.CountInstancesWithStatus(subscription, WorkflowStatus.NotSpecified);


                            ClientContext.ExecuteQueryRetry();

                            WriteVerbose(string.Format("Terminated => {0}", countTerminated.Value));
                            WriteVerbose(string.Format("Suspended => {0}", countSuspended.Value));
                            WriteVerbose(string.Format("Invalid => {0}", countInvalid.Value));
                            WriteVerbose(string.Format("Canceled => {0}", countCancelled.Value));
                            WriteVerbose(string.Format("Canceling => {0}", countCanceling.Value));
                            WriteVerbose(string.Format("Started => {0}", countStarted.Value));
                            WriteVerbose(string.Format("NotStarted => {0}", countNotStarted.Value));
                            WriteVerbose(string.Format("NotSpecified => {0}", countNotSpecified.Value));


                            var instances = deploymentService.Enumerate(subscription);
                            ClientContext.Load(instances);
                            ClientContext.ExecuteQueryRetry();
                            var wdx = 0;

                            while (instances != null && instances.AreItemsAvailable)
                            {
                                wdx += instances.Count;
                                foreach (var instance in instances)
                                {
                                    WriteVerbose(string.Format("Instance {0} => Status {1} => LastUpdated {2}", instance.Id, instance.Status, instance.LastUpdated));
                                }

                            }

                            WriteObject(instances);
                        }
                    }
                }
            }
            else
            {
                WriteObject(SelectedWeb.GetWorkflowInstances());
            }
        }
    }

}
