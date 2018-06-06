using System;
using System.Linq;
using System.Management.Automation;
using Microsoft.SharePoint.Client;
using InfrastructureAsCode.Powershell.PipeBinds;
using InfrastructureAsCode.Powershell.CmdLets;
using Microsoft.SharePoint.Client.WorkflowServices;
using System.Collections.Generic;

namespace InfrastructureAsCode.Powershell.Commands.Workflow
{
    [Cmdlet(VerbsLifecycle.Start, "IaCWorkflowInstance")]
    public class StartIaCWorkflowInstance : IaCCmdlet
    {
        [Parameter(Mandatory = false, HelpMessage = "The name of the workflow", Position = 0)]
        public string SubscriptionId { get; set; }

        [Parameter(Mandatory = false, HelpMessage = "A list to search the instances for", Position = 1)]
        public ListPipeBind List { get; set; }

        [Parameter(Mandatory = false)]
        public int ListItemId { get; set; }


        public override void ExecuteCmdlet()
        {

            var SelectedWeb = this.ClientContext.Web;


            var list = List.GetList(SelectedWeb);


            if (Guid.TryParse(SubscriptionId, out Guid subscriptionId))
            {

                var item = list.GetItemById("" + ListItemId);
                list.Context.Load(item, ictx => ictx.Id, ictx => ictx.ParentList.Id);
                list.Context.ExecuteQueryRetry();


                var instanceId = item.StartWorkflowInstance(subscriptionId, new Dictionary<string, object>());



                var instances = SelectedWeb.GetWorkflowInstances(item);
                foreach (var instance in instances)
                {
                    LogVerbose("Instance {0} => Status {1} => WF Status {2} => Created {3} => LastUpdated {4} => Subscription {5}", instance.Id, instance.Status, instance.UserStatus, instance.InstanceCreated, instance.LastUpdated, instance.WorkflowSubscriptionId);
                }
            }
        }
    }


}
