using System.Collections.Generic;

namespace InfrastructureAsCode.Core.Models
{
    public class SPGroupDefinitionModel
    {
        public SPGroupDefinitionModel()
        {
            this.ApprovalGroup = false;
            this.Users = new List<SPUserDefinitionModel>();
        }

        /// <summary>
        /// initialize group object
        /// </summary>
        /// <param name="title"></param>
        public SPGroupDefinitionModel(string title)
        {
            this.Title = title;
        }

        /// <summary>
        /// Stores the ID from the system where created
        /// </summary>
        public int Id { get; set; }

        public string Title { get; set; }

        public bool AllowMembersEditMembership { get; set; }

        public bool AllowRequestToJoinLeave { get; set; }

        public bool AutoAcceptRequestToJoinLeave { get; set; }

        public bool CanCurrentUserEditMembership { get; }

        public bool CanCurrentUserManageGroup { get; }

        public bool CanCurrentUserViewMembership { get; }

        public string Description { get; set; }

        public bool OnlyAllowMembersViewMembership { get; set; }

        public string RequestToJoinLeaveEmailSetting { get; set; }

        public bool ApprovalGroup { get; set; }

        /// <summary>
        /// A users collection for the group
        /// </summary>
        public ICollection<SPUserDefinitionModel> Users { get; set; }
    }
}