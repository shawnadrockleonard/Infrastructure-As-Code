using InfrastructureAsCode.Core;
using InfrastructureAsCode.Powershell.Commands.Base;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WorkflowServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;


namespace InfrastructureAsCode.Powershell.Commands.Workflow
{
    /// <summary>
    /// Queries the web for all subscriptions
    /// </summary>
    [Cmdlet(VerbsCommon.Get, "IaCWorkflows", SupportsShouldProcess = true)]
    public class GetIaCWorkflows : IaCCmdlet
    {
        #region Public Parameters 

        /// <summary>
        /// The workflow when specified will filter the subscriptions to ensure we only operate on a specific workflow
        /// </summary>
        [Parameter(Mandatory = false, HelpMessage = "A specific workflow name for only processing a specific workflow", Position = 0)]
        public string WorkflowName { get; set; }

        #endregion

        #region Private Variables

        #endregion


        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();

            var SelectedWeb = this.ClientContext.Web;

            var ilogger = new DefaultUsageLogger(
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

            IEnumerable<WorkflowSubscription> subscriptionCollection = null;

            if (!string.IsNullOrEmpty(WorkflowName))
            {
                var name = WorkflowName;
                var servicesManager = new WorkflowServicesManager(ClientContext, SelectedWeb);
                var subscriptionService = servicesManager.GetWorkflowSubscriptionService();
                var subscriptions = subscriptionService.EnumerateSubscriptions();
                subscriptionCollection = ClientContext.LoadQuery(from sub in subscriptions where sub.Name == name select sub);
                ClientContext.ExecuteQueryRetry();
            }
            else
            {
                var servicesManager = new WorkflowServicesManager(ClientContext, SelectedWeb);
                var subscriptionService = servicesManager.GetWorkflowSubscriptionService();
                var subscriptions = subscriptionService.EnumerateSubscriptions();
                ClientContext.Load(subscriptions);
                ClientContext.ExecuteQueryRetry();

                subscriptionCollection = subscriptions.ToArray();
            }

            // Log the workflow subscription
            foreach (var itemId in subscriptionCollection)
            {
                var msg = $"Workflow Subscription {itemId.Id} => Name {itemId.Name}";
                LogWarning(msg);
            }



            var workflowDefinitions = SelectedWeb.GetWorkflowDefinitions(false);
            foreach (var wfDefinition in workflowDefinitions)
            {
                var msg = $"Workflow Definition {wfDefinition.Id} => Name {wfDefinition.DisplayName}";
                LogWarning(msg);
            }


            var workflowInstances = SelectedWeb.GetWorkflowInstances();
            foreach (var wfInstance in workflowInstances)
            {
                var msg = $"Workflow Instance {wfInstance.Id} => Fault {wfInstance.FaultInfo}";
                LogWarning(msg);
            }


            var omservicesManager = new WorkflowServicesManager(ClientContext, SelectedWeb);
            var omsubscriptionService = omservicesManager.GetWorkflowDeploymentService();
            var omsubscriptions = omsubscriptionService.EnumerateDefinitions(false);
            ClientContext.Load(omsubscriptions);
            ClientContext.ExecuteQueryRetry();
            foreach(var omsub in omsubscriptions)
            {
                var msg = $"Workflow DS Def {omsub.Id} => List {omsub.AssociationUrl} => Name {omsub.DisplayName}";
                LogWarning(msg);
            }


            var queryAssociations = ClientContext.LoadQuery(SelectedWeb.WorkflowAssociations);
            ClientContext.ExecuteQueryRetry();
            foreach (var wfTemplate in queryAssociations)
            {
                var msg = $"Workflow Template {wfTemplate.Id} => List {wfTemplate.ListId} => Name {wfTemplate.Name}";
                LogWarning(msg);
            }


            var query = ClientContext.LoadQuery(SelectedWeb.WorkflowTemplates);
            ClientContext.ExecuteQueryRetry();
            foreach (var wfTemplate in query)
            {
                var msg = $"Workflow Template {wfTemplate.Id} => Declaritive {wfTemplate.IsDeclarative} => Name {wfTemplate.Name}";
                LogWarning(msg);
            }

        }
    }
}
