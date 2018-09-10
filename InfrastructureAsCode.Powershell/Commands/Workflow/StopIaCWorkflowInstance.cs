using InfrastructureAsCode.Core;
using InfrastructureAsCode.Core.Extensions;
using InfrastructureAsCode.Core.Models;
using InfrastructureAsCode.Powershell.Commands.Base;
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
    /// Queries a List/View and sets each list item to have its WF instance removed and eventually reset
    /// </summary>
    [Cmdlet(VerbsLifecycle.Stop, "IaCWorkflowInstance", SupportsShouldProcess = true)]
    public class StopIaCWorkflowInstance : IaCCmdlet
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

            var view = View.GetView(list);
            var internalViewFieldXml = new List<string>() { CAML.FieldRef("Id") };
            foreach (var vfield in view.ViewFields.Where(w => !w.Equals("ID", StringComparison.CurrentCultureIgnoreCase)))
            {
                internalViewFieldXml.Add(CAML.FieldRef(vfield));
            }

            var itemIds = new List<int>();
            var viewCaml = new CamlQuery()
            {
                ViewXml = CAML.ViewQuery(ViewScope.RecursiveAll, view.ViewQuery, string.Empty, CAML.ViewFields(internalViewFieldXml.ToArray()), view.RowLimit.ToString().ToInt32(5)),
                ListItemCollectionPosition = null
            };

            do
            {
                var items = list.GetItems(viewCaml);
                this.ClientContext.Load(items, ftx => ftx.ListItemCollectionPosition, ftx => ftx.Include(ftcx => ftcx.Id, ftcx => ftcx.ParentList.Id, ftcx => ftcx[FieldBoolean_RestartWorkflow]));
                this.ClientContext.ExecuteQueryRetry();
                viewCaml.ListItemCollectionPosition = items.ListItemCollectionPosition;

                foreach (var item in items)
                {
                    // Load ParentList ID to Pull Workflow Instances
                    var allinstances = SelectedWeb.GetWorkflowInstances(item);
                    if (allinstances.Any())
                    {
                        foreach (var instance in allinstances)
                        {
                            objects.Add(new SPWorkflowInstance(instance, item.Id));
                        }

                        itemIds.Add(item.Id);
                    }
                }

            }
            while (viewCaml.ListItemCollectionPosition != null);

            var rowprocessed = itemIds.Count();
            if (rowprocessed > 0 && this.ShouldProcess($"Setting Restart Flag for {rowprocessed} items"))
            {
                var rowdx = 0; var totaldx = rowprocessed;

                foreach (var itemId in itemIds)
                {
                    rowdx++;
                    totaldx--;

                    var wfItem = list.GetItemById(itemId);
                    list.Context.Load(wfItem);
                    wfItem[FieldBoolean_RestartWorkflow] = true;
                    wfItem.SystemUpdate();

                    if (rowdx >= 50 || totaldx <= 0)
                    {
                        list.Context.ExecuteQueryRetry();
                        Ilogger.LogInformation($"Processing {rowprocessed} rows; Persisted {rowdx} rows; {totaldx} remaining");
                        rowdx = 0;
                    }
                }

                var cancelDirty = false;
                var viewFieldXml = CAML.ViewFields(viewFields.Select(s => CAML.FieldRef(s)).ToArray());
                var terminatedWFCaml = new CamlQuery()
                {
                    ViewXml = CAML.ViewQuery(ViewScope.RecursiveAll, CAML.Where(CAML.Eq(CAML.FieldValue(FieldBoolean_RestartWorkflow, FieldType.Boolean.ToString("f"), 1.ToString()))), string.Empty, viewFieldXml, 100),
                    ListItemCollectionPosition = null
                };

                do
                {
                    var items = list.GetItems(terminatedWFCaml);
                    list.Context.Load(items, ftx => ftx.ListItemCollectionPosition, ftx => ftx.Include(ftcx => ftcx.Id, ftcx => ftcx.ParentList.Id, ftcx => ftcx[FieldBoolean_RestartWorkflow]));
                    list.Context.ExecuteQueryRetry();
                    terminatedWFCaml.ListItemCollectionPosition = items.ListItemCollectionPosition;

                    rowdx = 0; totaldx = items.Count();

                    foreach (var item in items)
                    {
                        rowdx++;
                        totaldx--;

                        var itemId = item.Id;
                        var allinstances = SelectedWeb.GetWorkflowInstances(item);
                        foreach (var instance in allinstances)
                        {
                            var instanceId = instance.Id;

                            var msg = $"List Item {itemId} => Cancelling subscription {instance.WorkflowSubscriptionId} instance Id {instanceId}";
                            Ilogger.LogWarning(msg);
                            cancelDirty = true;
                            instance.CancelWorkFlow();
                        }


                        if (cancelDirty
                            && (rowdx >= 50 || totaldx <= 0))
                        {
                            list.Context.ExecuteQueryRetry();
                            Ilogger.LogInformation($"Processing {rowprocessed} rows; Persisted {rowdx} rows; {totaldx} remaining");
                            rowdx = 0;
                        }
                    }

                }
                while (terminatedWFCaml.ListItemCollectionPosition != null);
            }


            WriteObject(objects, true);
        }
    }
}
