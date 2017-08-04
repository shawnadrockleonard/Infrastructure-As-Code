using InfrastructureAsCode.Powershell;
using InfrastructureAsCode.Powershell.CmdLets;
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Utilities;
using OfficeDevPnP.Core.Entities;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Powershell.Commands.Principals
{
    [Cmdlet(VerbsCommon.Remove, "IaCExternalUserFromSite")]
    [CmdletHelp("Removes external user from the sharepoint site and tenant collection", Category = "External")]
    public class RemoveIaCExternalUserFromSite : IaCAdminCmdlet
    {

        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
        public string SiteUrl;

        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 1)]
        public string UserName;

        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 2)]
        public string GroupName;

        [Parameter(Mandatory = false)]
        public SwitchParameter RemoveUser;


        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();
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

                    RemoveExternalUser( this.ClientContext, this.SiteUrl, this.UserName);

                }
            }
            catch (Exception ex)
            {
                LogError(ex, "Failed to execute RemoveUserProfileFromSite by site {0}", this.SiteUrl);
            }
        }


        private void RemoveExternalUser(ClientContext clientContext, string siteUrl, string profileName = null, int startIndex = 0, int pageSize = 50)
        {
            var invitedAs = profileName;
            if (profileName.IndexOf("#") > 0)
            {
                invitedAs = profileName.Substring(0, profileName.IndexOf("#"));
            }
            invitedAs = invitedAs.Replace("_", "@");

            var filterQuery = string.Format("InvitedAs -eq '{0}'", profileName);
            LogVerbose(string.Format("Removing External User with {0} at start {1} and page size {2}", profileName, startIndex, pageSize));
            var externalusers = this.OfficeTenant.GetExternalUsers(startIndex, pageSize, invitedAs, Microsoft.Online.SharePoint.TenantManagement.SortOrder.Ascending);
            var extCol = externalusers.ExternalUserCollection;
            clientContext.Load(externalusers);
            clientContext.Load(extCol);
            clientContext.ExecuteQuery();

            var userFound = extCol.Any(w => w.InvitedAs == invitedAs || w.InvitedAs.Contains(invitedAs));
            if (RemoveUser && userFound)
            {
                //10030000928913CB
                var externalUserId = extCol.FirstOrDefault(w => w.InvitedAs == invitedAs).UniqueId;
                var externalArray = new string[] { externalUserId };
                var externalResult = this.OfficeTenant.RemoveExternalUsers(externalArray);
                clientContext.ExecuteQuery();
            }

            if (externalusers.TotalUserCount > pageSize && extCol.Count == pageSize)
            {
                //# do some paging 
                startIndex += pageSize;
                RemoveExternalUser( clientContext, siteUrl, profileName, startIndex, pageSize);
            }
        }
    }
}
