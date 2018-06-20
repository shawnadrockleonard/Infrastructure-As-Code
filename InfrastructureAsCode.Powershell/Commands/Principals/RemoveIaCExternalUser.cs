using InfrastructureAsCode.Core;
using InfrastructureAsCode.Core.Extensions;
using InfrastructureAsCode.Core.Models;
using InfrastructureAsCode.Powershell.CmdLets;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;

namespace InfrastructureAsCode.Powershell.Commands.Principals
{
    /// <summary>
    /// Removes a user from the sharepoint site
    /// </summary>
    [Cmdlet(VerbsCommon.Remove, "IaCExternalUser", SupportsShouldProcess = true)]
    public class RemoveIaCExternalUser : IaCAdminCmdlet
    {

        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
        public string UserName { get; set; }


        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();
            try
            {
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


                var externalUsers = this.ClientContext.CheckExternalUser(ilogger, invitedAs);
                foreach (var externalUsr in externalUsers)
                {

                    RemoveExternalUser(externalUsers, invitedAs);
                }
            }
            catch (Exception ex)
            {
                LogError(ex, "Failed to execute RemoveUserProfileFromSite -UserName {0}", this.UserName);
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
                    var externalResult = OfficeTenantContext.RemoveExternalUsers(externalArray);
                    OfficeTenantContext.Context.ExecuteQueryRetry();
                }
            }

        }
    }
}
