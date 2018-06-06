using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using InfrastructureAsCode.Core.Extensions;
using Microsoft.SharePoint.Client;
using System.Management.Automation;
using InfrastructureAsCode.Core.Models;
using OfficeDevPnP.Core.Utilities;
using Microsoft.SharePoint.Client.WorkflowServices;
using InfrastructureAsCode.Powershell.CmdLets;
using InfrastructureAsCode.Powershell.PipeBinds;
using InfrastructureAsCode.Core.Reports;


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

            WorkflowSubscriptionCollection subscriptions = null;

            if (!string.IsNullOrEmpty(WorkflowName))
            {
                var servicesManager = new WorkflowServicesManager(ClientContext, SelectedWeb);
                var subscriptionService = servicesManager.GetWorkflowSubscriptionService();
                subscriptions = subscriptionService.EnumerateSubscriptions();
                ClientContext.Load(subscriptions, subs => subs.Where(sub => sub.Name == WorkflowName));
                ClientContext.ExecuteQueryRetry();


            }
            else
            {
                var servicesManager = new WorkflowServicesManager(ClientContext, SelectedWeb);
                var subscriptionService = servicesManager.GetWorkflowSubscriptionService();
                subscriptions = subscriptionService.EnumerateSubscriptions();
                ClientContext.Load(subscriptions);
                ClientContext.ExecuteQueryRetry();

            }

            // Log the workflow subscription
            foreach (var itemId in subscriptions)
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
