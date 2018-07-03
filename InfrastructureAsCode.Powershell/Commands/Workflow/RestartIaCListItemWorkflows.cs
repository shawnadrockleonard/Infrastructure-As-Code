using InfrastructureAsCode.Core;
using InfrastructureAsCode.Core.Extensions;
using InfrastructureAsCode.Core.Models;
using InfrastructureAsCode.Powershell.CmdLets;
using InfrastructureAsCode.Powershell.PipeBinds;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WorkflowServices;
using OfficeDevPnP.Core.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;

namespace InfrastructureAsCode.Powershell.Commands.Workflow
{
    /// <summary>
    /// Enables the ease of restarting workflow in the specified list
    /// </summary>
    [Cmdlet(VerbsLifecycle.Restart, "IaCListItemWorkflows", SupportsShouldProcess = true)]
    public class RestartIaCListItemWorkflows : IaCCmdlet
    {
        #region Public Parameters 

        /// <summary>
        /// A list to search the instances for which workflow instances will be restarted
        /// </summary>
        [Parameter(Mandatory = true, HelpMessage = "A list to search the instances for", Position = 0)]
        public ListPipeBind List { get; set; }

        /// <summary>
        /// The workflow when specified will filter the subscriptions to ensure we only operate on a specific workflow
        /// </summary>
        [Parameter(Mandatory = false, HelpMessage = "A specific workflow name for only processing a specific workflow", Position = 1)]
        public string WorkflowName { get; set; }


        [Parameter(Mandatory = false, HelpMessage = "A specific workflow name for only processing a specific workflow", Position = 2)]
        public string WorkflowColumnName { get; set; }

        #endregion

        #region Private Variables

        internal const string FieldBoolean_RestartWorkflow = "RestartWorkflow";

        private ITraceLogger ilogger { get; set; }

        #endregion


        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();

            var SelectedWeb = this.ClientContext.Web;

            ilogger = new DefaultUsageLogger(
                (string msg, object[] margs) =>
                {
                    LogDebugging(msg, margs);
                },
                (string msg, object[] margs) =>
                {
                    LogWarning(msg, margs);
                },
                (Exception ex, string msg, object[] margs) =>
                {
                    LogError(ex, msg, margs);
                });

            var list = List.GetList(SelectedWeb);
            SelectedWeb.Context.Load(list, lctx => lctx.Id);
            SelectedWeb.Context.ExecuteQueryRetry();

            var workflowSubscriptionId = default(Nullable<Guid>);
            if (!string.IsNullOrEmpty(WorkflowName))
            {
                var servicesManager = new WorkflowServicesManager(ClientContext, SelectedWeb);
                var subscriptionService = servicesManager.GetWorkflowSubscriptionService();
                var subscriptions = subscriptionService.EnumerateSubscriptionsByList(list.Id);
                //foreach (WorkflowSubscription subs1 in subscriptions)
                //{
                //    if (subs1.Name == WorkflowName) { ClientContext.Load(subs1); }
                //}
                ClientContext.Load(subscriptions);
                ClientContext.ExecuteQueryRetry();
                if (subscriptions.Any(subs => subs.Name == WorkflowName))
                {
                    var subscription = subscriptions.FirstOrDefault(subs => subs.Name == WorkflowName);
                    workflowSubscriptionId = subscription.Id;
                }
            }

            // Check if the field exists
            var viewFields = new string[] { "Id", "Title", WorkflowColumnName };
            var viewFieldXml = CAML.ViewFields(viewFields.Select(s => CAML.FieldRef(s)).ToArray());
            var internalFields = new List<string>();
            internalFields.AddRange(viewFields);

            try
            {
                var checkFields = list.GetFields(internalFields.ToArray());
            }
            catch (Exception ex)
            {
                LogError(ex, $"Failed to retreive the fields {string.Join(";", viewFields)} with Msg {ex.Message}");
                return;
            }


            var itemIds = new List<int>();
            var viewCaml = new CamlQuery()
            {
                ViewXml = CAML.ViewQuery(
                ViewScope.RecursiveAll,
                CAML.Where(CAML.Neq(CAML.FieldValue(WorkflowColumnName, FieldType.WorkflowStatus.ToString("f"), WorkflowStatus.Completed.ToString("D")))),
                string.Empty,
                viewFieldXml,
                100)
            };
            ListItemCollectionPosition itemPosition = null;

            while (true)
            {
                viewCaml.ListItemCollectionPosition = itemPosition;
                var items = list.GetItems(viewCaml);
                list.Context.Load(items);
                list.Context.ExecuteQueryRetry();
                itemPosition = items.ListItemCollectionPosition;

                foreach (var item in items)
                {
                    itemIds.Add(item.Id);
                }

                if (itemPosition == null)
                {
                    break;
                }
            }

            // Workflow status to re-start!
            var workflowStati = new List<Microsoft.SharePoint.Client.WorkflowServices.WorkflowStatus>()
            {
                Microsoft.SharePoint.Client.WorkflowServices.WorkflowStatus.Started,
                Microsoft.SharePoint.Client.WorkflowServices.WorkflowStatus.Suspended,
                Microsoft.SharePoint.Client.WorkflowServices.WorkflowStatus.Invalid
            };
            foreach (var itemId in itemIds)
            {
                // Retreive the ListItem
                var item = list.GetItemById("" + itemId);
                list.Context.Load(item, ictx => ictx.Id, ictx => ictx.ParentList.Id);
                list.Context.ExecuteQueryRetry();

                // Variables for processing
                var subscriptionIds = new List<Guid>();
                var allinstances = SelectedWeb.GetWorkflowInstances(item);
                var terminationInstances = allinstances.Where(instance => workflowStati.Any(ws => ws == instance.Status)
                        && (!workflowSubscriptionId.HasValue || (!workflowSubscriptionId.HasValue && instance.WorkflowSubscriptionId == workflowSubscriptionId))).ToList();

                // Cancel the existing failed workflow instances
                foreach (var instance in terminationInstances)
                {
                    var instanceId = instance.Id;
                    subscriptionIds.Add(instance.WorkflowSubscriptionId);

                    var msg = string.Format("List Item {0} => Cancelling subscription {1} instance Id {2}", itemId, instance.WorkflowSubscriptionId, instanceId);
                    if (ShouldProcess(msg))
                    {
                        instance.CancelWorkFlow();
                        LogWarning("List Item {0} => Cancelled subscription {1} instance Id {2}", itemId, instance.WorkflowSubscriptionId, instanceId);
                    }
                }

                // Instantiate the workflow subscription
                foreach (var subscriptionId in subscriptionIds)
                {
                    var msg = string.Format("List Item {0} => Start workflow subscription id {1}", itemId, subscriptionId);
                    if (ShouldProcess(msg))
                    {
                        var instanceId = item.StartWorkflowInstance(subscriptionId, new Dictionary<string, object>());
                        LogWarning("List Item {0} => Successfully restarted subscription {1} with new instance Id {2}", itemId, subscriptionId, instanceId);
                    }
                }

                // Retreive the Item workflow instances and print to the console
                var instances = SelectedWeb.GetWorkflowInstances(item);
                foreach (var instance in instances)
                {
                    LogVerbose("List Item {0} => Instance {1} => Status {2} => WF Status {3} => Created {4} => LastUpdated {5} => Subscription {6}", itemId, instance.Id, instance.Status, instance.UserStatus, instance.InstanceCreated, instance.LastUpdated, instance.WorkflowSubscriptionId);
                }
            }
        }
    }
}
