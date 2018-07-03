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
    [Cmdlet(VerbsLifecycle.Start, "IaCWorkflowInstance")]
    public class StartIaCWorkflowInstance : IaCCmdlet
    {

        #region Public Parameters 

        /// <summary>
        /// A list to search the instances for which workflow instances will be restarted
        /// </summary>
        [Parameter(Mandatory = true, HelpMessage = "A list to search the instances for")]
        public ListPipeBind List { get; set; }

        /// <summary>
        /// A view to search the instances for which workflow instances will be restarted
        /// </summary>
        [Parameter(Mandatory = true, HelpMessage = "A view to search the instances for", Position = 0)]
        public ViewPipeBind View { get; set; }

        /// <summary>
        /// The workflow when specified will filter the subscriptions to ensure we only operate on a specific workflow
        /// </summary>
        [Parameter(Mandatory = true, HelpMessage = "A specific workflow name for only processing a specific workflow")]
        public string WorkflowName { get; set; }

        #endregion


        #region Private Variables

        internal const string FieldBoolean_RestartWorkflow = "RestartWorkflow";

        private ITraceLogger Ilogger { get; set; }

        #endregion


        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();


            var objects = new List<SPWorkflowInstance>();


            var FieldDefinitions = new List<SPFieldDefinitionModel>
            {
                new SPFieldDefinitionModel(FieldType.Boolean)
                {
                    FieldGuid = new Guid("da2872c4-e9b6-4804-9837-6e9dd85ecd7e"),
                    InternalName = FieldBoolean_RestartWorkflow,
                    Description = "RestartWorkflow provides a way to identify items that should be restarted.",
                    Title = FieldBoolean_RestartWorkflow,
                    MaxLength = 255,
                    DefaultValue = "No"
                }
            };


            var SelectedWeb = this.ClientContext.Web;

            Ilogger = new DefaultUsageLogger(
                (string msg, object[] margs) =>
                {
                    LogVerbose(msg, margs);
                },
                (string msg, object[] margs) =>
                {
                    LogWarning(msg, margs);
                },
                (Exception ex, string msg, object[] margs) =>
                {
                    LogError(ex, msg, margs);
                });

            var list = List.GetList(SelectedWeb, lctx => lctx.Id, lctx => lctx.Title);

            var workflowSubscriptionId = new List<Guid>();
            if (!string.IsNullOrEmpty(WorkflowName))
            {
                var servicesManager = new WorkflowServicesManager(ClientContext, SelectedWeb);
                var subscriptionService = servicesManager.GetWorkflowSubscriptionService();
                var subscriptions = subscriptionService.EnumerateSubscriptionsByList(list.Id);

                ClientContext.Load(subscriptions);
                ClientContext.ExecuteQueryRetry();

                subscriptions.Where(subs => subs.Name == WorkflowName).Select(s => s.Id).ToList().ForEach(subscriptionId =>
                {
                    workflowSubscriptionId.Add(subscriptionId);
                });
            }


            // Check if the field exists
            var viewFields = new string[] { "Id", "Title", FieldBoolean_RestartWorkflow };
            var internalFields = new List<string>();
            internalFields.AddRange(FieldDefinitions.Select(s => s.InternalName));

            try
            {
                var checkFields = list.GetFields(internalFields.ToArray());
            }
            catch (Exception ex)
            {
                Ilogger.LogError(ex, "Failed to retreive the fields {0}", ex.Message);

                foreach (var field in FieldDefinitions)
                {
                    // provision the column
                    var provisionedColumn = list.CreateListColumn(field, Ilogger, null);
                    if (provisionedColumn != null)
                    {
                        internalFields.Add(provisionedColumn.InternalName);
                    }
                }
            }

            var workflowStati = new List<Microsoft.SharePoint.Client.WorkflowServices.WorkflowStatus>()
            {
                Microsoft.SharePoint.Client.WorkflowServices.WorkflowStatus.Started
            };

            var rowprocessed = 0;
            var rowdx = 0; var totaldx = 0;
            var startedDirty = false;
            var viewFieldXml = CAML.ViewFields(viewFields.Select(s => CAML.FieldRef(s)).ToArray());
            var initiateWFCaml = new CamlQuery()
            {
                ViewXml = CAML.ViewQuery(ViewScope.RecursiveAll, CAML.Where(CAML.Eq(CAML.FieldValue(FieldBoolean_RestartWorkflow, FieldType.Boolean.ToString("f"), 1.ToString()))), string.Empty, viewFieldXml, 100),
                ListItemCollectionPosition = null
            };


            do
            {
                var items = list.GetItems(initiateWFCaml);
                list.Context.Load(items, ftx => ftx.ListItemCollectionPosition, ftx => ftx.Include(ftcx => ftcx.Id, ftcx => ftcx.ParentList.Id, ftcx => ftcx[FieldBoolean_RestartWorkflow]));
                list.Context.ExecuteQueryRetry();
                initiateWFCaml.ListItemCollectionPosition = items.ListItemCollectionPosition;

                rowprocessed = items.Count();
                rowdx = 0; totaldx = rowprocessed;

                foreach (var item in items)
                {
                    rowdx++;
                    totaldx--;

                    var itemId = item.Id;
                    var allinstances = SelectedWeb.GetWorkflowInstances(item);
                    if (!allinstances.Any(w => workflowStati.Any(wx => wx == w.Status)))
                    {
                        foreach (var subscriptionId in workflowSubscriptionId)
                        {
                            Ilogger.LogWarning($"List Item {itemId} => Restarting subscription {subscriptionId}");
                            var instanceId = item.StartWorkflowInstance(subscriptionId, new Dictionary<string, object>());
                            Ilogger.LogWarning($"List Item {itemId} => Successfully restarted subscription {subscriptionId} with new instance Id {instanceId}");
                            startedDirty = true;

                            objects.Add(new SPWorkflowInstance()
                            {
                                Id = instanceId,
                                WorkflowSubscriptionId = subscriptionId,
                                ListItemId = itemId,
                                InstanceCreated = DateTime.UtcNow
                            });
                        }

                        var wfItem = list.GetItemById(itemId);
                        list.Context.Load(wfItem);
                        wfItem[FieldBoolean_RestartWorkflow] = false;
                        wfItem.SystemUpdate();
                    }

                    if (startedDirty && (rowdx >= 50 || totaldx <= 0))
                    {
                        list.Context.ExecuteQueryRetry();
                        Ilogger.LogInformation($"Processing {rowprocessed} rows; Persisted {rowdx} rows; {totaldx} remaining");
                        rowdx = 0;
                    }
                }

            }
            while (initiateWFCaml.ListItemCollectionPosition != null);


            WriteObject(objects, true);
        }

    }
}
