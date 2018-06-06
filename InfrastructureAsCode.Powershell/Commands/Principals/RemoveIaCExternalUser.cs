using InfrastructureAsCode.Powershell;
using InfrastructureAsCode.Powershell.CmdLets;
using InfrastructureAsCode.Core.Utilities;
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
    /// <summary>
    /// Removes a user from the sharepoint site
    /// </summary>
    [Cmdlet(VerbsCommon.Remove, "IaCExternalUser", SupportsShouldProcess = true)]
    public class RemoveIaCExternalUser : IaCAdminCmdlet
    {

        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
        public string SiteUrl { get; set; }

        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 1)]
        public string UserName { get; set; }


        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            if (string.IsNullOrEmpty(this.SiteUrl))
            {
                throw new InvalidOperationException("Site is not valid.");
            }
        }

        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();
            try
            {
                RemoveExternalUser(this.SiteUrl, this.UserName);
            }
            catch (Exception ex)
            {
                LogError(ex, "Failed to execute RemoveUserProfileFromSite by site {0}", this.SiteUrl);
            }
        }


        private void RemoveExternalUser(string SiteUrl, string UserName = null, int startIndex = 0, int pageSize = 50)
        {
            var invitedAs = UserName;
            if (UserName.IndexOf("#") > 0)
            {
                invitedAs = UserName.Substring(0, UserName.IndexOf("#"));
            }
            //invitedAs = invitedAs.Replace("_", "@");

            var filterQuery = string.Format("InvitedAs -eq '{0}'", UserName);
            LogVerbose(string.Format("Removing External User with {0} at start {1} and page size {2}", UserName, startIndex, pageSize));
            var externalusers = OfficeTenantContext.GetExternalUsers(startIndex, pageSize, invitedAs, Microsoft.Online.SharePoint.TenantManagement.SortOrder.Ascending);
            var extCol = externalusers.ExternalUserCollection;
            OfficeTenantContext.Context.Load(externalusers);
            OfficeTenantContext.Context.Load(extCol);
            OfficeTenantContext.Context.ExecuteQueryRetry();

            var userFound = extCol.Any(w => w.InvitedAs == invitedAs || w.InvitedAs.Contains(invitedAs));
            if (userFound)
            {
                var user = extCol.FirstOrDefault(w => w.InvitedAs == invitedAs || w.InvitedAs.Contains(invitedAs));
                LogVerbose("User {0} invited by {1} accepted as {2}", invitedAs, user.InvitedBy, user.AcceptedAs);
                if (this.ShouldProcess(string.Format("User {0} will be removed from the tenant {1}", invitedAs, OfficeTenantContext.Context.Url)))
                {
                    //10030000928913CB
                    var externalUserId = extCol.FirstOrDefault(w => w.InvitedAs == invitedAs).UniqueId;
                    var externalArray = new string[] { externalUserId };
                    var externalResult = OfficeTenantContext.RemoveExternalUsers(externalArray);
                    OfficeTenantContext.Context.ExecuteQueryRetry();
                }
            }

            if (externalusers.TotalUserCount > pageSize && extCol.Count == pageSize)
            {
                //# do some paging 
                startIndex += pageSize;
                RemoveExternalUser(SiteUrl, UserName, startIndex, pageSize);
            }
        }
    }
}
