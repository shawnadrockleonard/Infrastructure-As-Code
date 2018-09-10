using InfrastructureAsCode.Core;
using InfrastructureAsCode.Core.Extensions;
using InfrastructureAsCode.Core.Models;
using InfrastructureAsCode.Powershell.Commands.Base;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;

namespace InfrastructureAsCode.Powershell.Commands.Principals
{
    [Cmdlet(VerbsCommon.Remove, "IaCExternalUserFromSite", SupportsShouldProcess = true)]
    [CmdletHelp("Removes external user from the sharepoint site and tenant collection", Category = "External")]
    public class RemoveIaCExternalUserFromSite : IaCAdminCmdlet
    {

        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
        public string SiteUrl { get; set; }

        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 1)]
        public string UserName { get; set; }

        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 2)]
        public string GroupName { get; set; }



        protected override void OnBeginInitialize()
        {
            base.OnBeginInitialize();

            if (string.IsNullOrEmpty(this.SiteUrl))
            {
                throw new InvalidOperationException("Site is not valid.");
            }
        }


        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();

            var ilogger = new DefaultUsageLogger(LogVerbose, LogWarning, LogError);
            var invitedAs = string.Empty;

            if (!string.IsNullOrEmpty(this.UserName))
            {
                invitedAs = this.UserName.Replace($"{ClaimIdentifier}|", string.Empty).Trim();
                if (invitedAs.IndexOf("#") > 0)
                {
                    var loginIdentity = invitedAs.IndexOf("#");
                    invitedAs = invitedAs.Substring(0, loginIdentity);
                }
                invitedAs = invitedAs.Replace("_", "@");
            }


            try
            {
                using (var thisContext = this.ClientContext.Clone(this.SiteUrl))
                {
                    thisContext.Credentials = this.ClientContext.Credentials;

                    var qroups = thisContext.Web.SiteGroups;
                    thisContext.Load(qroups);

                    var groupName = qroups.GetByName(GroupName);
                    thisContext.Load(groupName, grp => grp.LoginName, grp => grp.Id, grp => grp.Users);

                    var users = groupName.Users;
                    thisContext.Load(users);
                    thisContext.ExecuteQuery();

                    var userInGroup = users.FirstOrDefault(u => u.LoginName == this.UserName || u.Email == this.UserName);
                    if (userInGroup != null)
                    {
                        groupName.Users.RemoveById(userInGroup.Id);
                        thisContext.ExecuteQuery();
                    }



                    var externalUsers = this.ClientContext.CheckExternalUser(ilogger, invitedAs);

                    RemoveExternalUser(externalUsers, invitedAs);

                }
            }
            catch (Exception ex)
            {
                LogError(ex, "Failed to execute RemoveUserProfileFromSite by site {0}", this.SiteUrl);
            }
        }


        private void RemoveExternalUser(IList<SPExternalUserEntity> users, string invitedAs)
        {
            LogVerbose($"Removing External User with {invitedAs}");

            foreach (var user in users.Where(w => w.InvitedAs == invitedAs || w.InvitedAs.Contains(invitedAs)))
            {
                LogVerbose("User {0} invited by {1} accepted as {2}", invitedAs, user.InvitedBy, user.AcceptedAs);

                if (this.ShouldProcess(string.Format("User {0} will be removed from the tenant {1}", invitedAs, OfficeTenantContext.Context.Url)))
                {
                    //10030000928913CB
                    var externalUserId = user.UniqueId;
                    var externalArray = new string[] { externalUserId };
                    var externalResult = this.OfficeTenantContext.RemoveExternalUsers(externalArray);
                    this.ClientContext.ExecuteQuery();
                }
            }
        }
    }
}
