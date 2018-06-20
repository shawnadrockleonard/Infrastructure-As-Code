using Microsoft.SharePoint.Client.WorkflowServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Models
{
    /// <summary>
    /// A workflow instance
    ///     If ListItemId is not null then it is a ListItem workflow instance
    /// </summary>
    public class SPWorkflowInstance
    {
        public SPWorkflowInstance() { }

        public SPWorkflowInstance(WorkflowInstance instance, Nullable<int> listItemId = default(Nullable<int>))
        {
            this.FaultInfo = instance.FaultInfo;
            this.Id = instance.Id;
            this.InstanceCreated = instance.InstanceCreated;
            this.LastUpdated = instance.LastUpdated;
            this.Status = instance.Status;
            this.UserStatus = instance.UserStatus;
            this.WorkflowSubscriptionId = instance.WorkflowSubscriptionId;
            this.ListItemId = listItemId;
        }

        public string FaultInfo { get; set; }

        public Guid Id { get; set; }

        public DateTime InstanceCreated { get; set; }

        public DateTime LastUpdated { get; set; }


        public WorkflowStatus Status { get; set; }

        public string UserStatus { get; set; }

        public Guid WorkflowSubscriptionId { get; set; }


        public Nullable<int> ListItemId { get; set; }

        public override string ToString()
        {
            return $"Instance {Id} => Status {Status.ToString("f")} => WF Status {UserStatus} => Created {InstanceCreated} => LastUpdated {LastUpdated} => Subscription {WorkflowSubscriptionId}";
        }
    }
}
