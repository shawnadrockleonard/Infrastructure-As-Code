using System.Management.Automation;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WorkflowServices;
using InfrastructureAsCode.Powershell.PipeBinds;
using InfrastructureAsCode.Powershell.CmdLets;
using System;

namespace InfrastructureAsCode.Powershell.Commands.Workflow
{
    /// <summary>
    /// https://msdn.microsoft.com/en-us/library/office/dn481315.aspx
    /// </summary>
    [Cmdlet(VerbsCommon.Get, "IaCWorkflowInstances")]
    public class GetIaCWorkflowInstances : IaCCmdlet
    {
        [Parameter(Mandatory = false, HelpMessage = "The name of the workflow", Position = 0)]
        public string Name { get; set; }

        [Parameter(Mandatory = false, HelpMessage = "A list to search the instances for", Position = 1)]
        public ListPipeBind List { get; set; }

        [Parameter(Mandatory = false)]
        public Nullable<int> ListItemId { get; set; }

        [Parameter(Mandatory = false)]
        public SwitchParameter PublishedOnly = true;

        public override void ExecuteCmdlet()
        {
            var SelectedWeb = this.ClientContext.Web;

            if (List != null)
            {
                var list = List.GetList(SelectedWeb);

                if (ListItemId.HasValue)
                {
                    var item = list.GetItemById("" + ListItemId);
                    list.Context.Load(item, ictx => ictx.Id, ictx => ictx.ParentList.Id);
                    list.Context.ExecuteQueryRetry();

                    var instances = SelectedWeb.GetWorkflowInstances(item);
                    foreach (var instance in instances)
                    {
                        LogVerbose("Instance {0} => Status {1} => WF Status {2} => Created {3} => LastUpdated {4} => Subscription {5}", instance.Id, instance.Status, instance.UserStatus, instance.InstanceCreated, instance.LastUpdated, instance.WorkflowSubscriptionId);
                    }
                    
                }
                else
                {
                    if (!string.IsNullOrEmpty(Name))
                    {
                        var servicesManager = new WorkflowServicesManager(ClientContext, SelectedWeb);
                        var deploymentService = servicesManager.GetWorkflowInstanceService();

                        WorkflowSubscription workflowSubscription = list.GetWorkflowSubscription(Name);
                        WriteSubscriptionInstances(deploymentService, workflowSubscription);

                    }
                    else
                    {
                        var servicesManager = new WorkflowServicesManager(ClientContext, SelectedWeb);
                        var deploymentService = servicesManager.GetWorkflowInstanceService();

                        var subscriptionService = servicesManager.GetWorkflowSubscriptionService();
                        var subscriptions = subscriptionService.EnumerateSubscriptionsByList(list.Id);
                        ClientContext.Load(subscriptions);
                        ClientContext.ExecuteQueryRetry();

                        foreach (var subscription in subscriptions)
                        {
                            WriteSubscriptionInstances(deploymentService, subscription);
                        }
                    }
                }
            }
            else
            {
                WriteObject(SelectedWeb.GetWorkflowInstances());
            }
        }

        private void WriteSubscriptionInstances(WorkflowInstanceService deploymentService, WorkflowSubscription workflowSubscription)
        {
            if (workflowSubscription != null && !workflowSubscription.ServerObjectIsNull())
            {
                var countTerminated = deploymentService.CountInstancesWithStatus(workflowSubscription, WorkflowStatus.Terminated);
                var countSuspended = deploymentService.CountInstancesWithStatus(workflowSubscription, WorkflowStatus.Suspended);
                var countInvalid = deploymentService.CountInstancesWithStatus(workflowSubscription, WorkflowStatus.Invalid);
                var countCancelled = deploymentService.CountInstancesWithStatus(workflowSubscription, WorkflowStatus.Canceled);
                var countCanceling = deploymentService.CountInstancesWithStatus(workflowSubscription, WorkflowStatus.Canceling);
                var countStarted = deploymentService.CountInstancesWithStatus(workflowSubscription, WorkflowStatus.Started);
                var countNotStarted = deploymentService.CountInstancesWithStatus(workflowSubscription, WorkflowStatus.NotStarted);
                var countNotSpecified = deploymentService.CountInstancesWithStatus(workflowSubscription, WorkflowStatus.NotSpecified);


                ClientContext.ExecuteQueryRetry();

                LogVerbose("Terminated => {0}", countTerminated.Value);
                LogVerbose("Suspended => {0}", countSuspended.Value);
                LogVerbose("Invalid => {0}", countInvalid.Value);
                LogVerbose("Canceled => {0}", countCancelled.Value);
                LogVerbose("Canceling => {0}", countCanceling.Value);
                LogVerbose("Started => {0}", countStarted.Value);
                LogVerbose("NotStarted => {0}", countNotStarted.Value);
                LogVerbose("NotSpecified => {0}", countNotSpecified.Value);


                var instances = deploymentService.Enumerate(workflowSubscription);
                ClientContext.Load(instances);
                ClientContext.ExecuteQueryRetry();
                var wdx = instances.Count;

                foreach (var instance in instances)
                {
                    LogVerbose("Instance {0} => Status {1} => WF Status {2} => Created {3} => LastUpdated {4} => Subscription {5}", instance.Id, instance.Status, instance.UserStatus, instance.InstanceCreated, instance.LastUpdated, instance.WorkflowSubscriptionId);
                }

            }
        }
    }

}
