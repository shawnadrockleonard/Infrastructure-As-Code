using System.Management.Automation;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WorkflowServices;
using InfrastructureAsCode.Powershell.PipeBinds;
using InfrastructureAsCode.Powershell.Commands.Base;
using System;
using System.Collections;
using System.Linq;
using InfrastructureAsCode.Core.Models;
using System.Collections.Generic;
using OfficeDevPnP.Core.Utilities;

namespace InfrastructureAsCode.Powershell.Commands.Workflow
{
    /// <summary>
    /// https://msdn.microsoft.com/en-us/library/office/dn481315.aspx
    /// </summary>
    [Cmdlet(VerbsCommon.Get, "IaCWorkflowInstances")]
    public class GetIaCWorkflowInstances : IaCCmdlet
    {

        #region Public Parameters 

        [Parameter(Mandatory = false, HelpMessage = "A list to search the instances for", Position = 0)]
        public ListPipeBind List { get; set; }


        [Parameter(Mandatory = false, HelpMessage = "The name of the workflow", Position = 1)]
        public string WorkflowName { get; set; }


        [Parameter(Mandatory = false)]
        public SwitchParameter DeepScan { get; set; }

        #endregion


        internal IList<SPWorkflowInstance> Instances { get; private set; }


        public override void ExecuteCmdlet()
        {
            var SelectedWeb = this.ClientContext.Web;


            var list = List.GetList(SelectedWeb);


            if (!string.IsNullOrEmpty(WorkflowName))
            {
                var servicesManager = new WorkflowServicesManager(ClientContext, SelectedWeb);
                var deploymentService = servicesManager.GetWorkflowInstanceService();

                WorkflowSubscription workflowSubscription = list.GetWorkflowSubscription(WorkflowName);
                WriteSubscriptionInstances(list, deploymentService, workflowSubscription);

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
                    WriteSubscriptionInstances(list, deploymentService, subscription);
                }
            }


            WriteObject(Instances);
        }

        private void WriteSubscriptionInstances(List list, WorkflowInstanceService deploymentService, WorkflowSubscription workflowSubscription)
        {
            Instances = new List<SPWorkflowInstance>();

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


                if (!DeepScan)
                {
                    var instances = deploymentService.Enumerate(workflowSubscription);
                    ClientContext.Load(instances);
                    ClientContext.ExecuteQueryRetry();

                    LogVerbose($"Instance {instances.Count}...");

                    foreach (var instance in instances)
                    {
                        Instances.Add(new SPWorkflowInstance(instance));
                    }

                }
                else
                {
                    var idx = 1;
                    var viewCaml = new CamlQuery()
                    {
                        ViewXml = CAML.ViewQuery(string.Empty, string.Empty, 100),
                        ListItemCollectionPosition = null
                    };

                    do
                    {
                        LogVerbose($"Deep search itr=>{idx++} paging => {viewCaml.ListItemCollectionPosition?.PagingInfo}");
                        var items = list.GetItems(viewCaml);
                        this.ClientContext.Load(items, ftx => ftx.ListItemCollectionPosition, ftx => ftx.Include(ftcx => ftcx.Id, ftcx => ftcx.ParentList.Id));
                        this.ClientContext.ExecuteQueryRetry();
                        viewCaml.ListItemCollectionPosition = items.ListItemCollectionPosition;

                        foreach (var item in items)
                        {
                            // Load ParentList ID to Pull Workflow Instances
                            var allinstances = ClientContext.Web.GetWorkflowInstances(item);
                            if (allinstances.Any())
                            {
                                foreach (var instance in allinstances)
                                {
                                    Instances.Add(new SPWorkflowInstance(instance, item.Id));
                                }
                            }
                        }

                    }
                    while (viewCaml.ListItemCollectionPosition != null);
                }

            }
        }

    }

}
