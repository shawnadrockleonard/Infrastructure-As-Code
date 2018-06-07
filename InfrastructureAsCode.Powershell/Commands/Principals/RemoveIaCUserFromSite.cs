using InfrastructureAsCode.Powershell.CmdLets;
using Microsoft.SharePoint.Client;
using System;
using System.Linq;
using System.Management.Automation;

namespace InfrastructureAsCode.Powershell.Commands.Principals
{
    /// <summary>
    /// Removes a user from the sharepoint site
    /// </summary>
    [Cmdlet(VerbsCommon.Remove, "IaCUserFromSite", SupportsShouldProcess = true)]
    public class RemoveIaCUserFromSite : IaCCmdlet
    {

        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 1)]
        public string UserName { get; set; }


        [Parameter(Mandatory = false, ValueFromPipeline = true, Position = 2)]
        public string GroupName { get; set; }



        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();
            try
            {
                var qusers = this.ClientContext.Web.SiteUsers;
                this.ClientContext.Load(qusers, iss => iss.Include(uiss => uiss.Id, uiss => uiss.Email, uiss => uiss.Title, uiss => uiss.LoginName, uiss => uiss.Groups));
                this.ClientContext.ExecuteQueryRetry();

                var quser = qusers.FirstOrDefault(f => f.Email.Equals(this.UserName, StringComparison.CurrentCultureIgnoreCase));
                if (quser != null)
                {
                    LogVerbose("User {0} found in site {1}; Enumerating Groups now", this.UserName, this.ClientContext.Url);
                    var groupsFound = quser.Groups.Count();

                    foreach (var qgroup in quser.Groups.Where(w =>
                    (!string.IsNullOrEmpty(this.GroupName) && (w.Title.Equals(this.GroupName, StringComparison.CurrentCultureIgnoreCase) || w.LoginName.Equals(this.GroupName, StringComparison.CurrentCultureIgnoreCase)))
                    || (string.IsNullOrEmpty(this.GroupName))))
                    {
                        if (this.ShouldProcess(string.Format("Sharepoint user {0} will be removed from {1}", this.UserName, qgroup.Title)))
                        {
                            quser.Groups.RemoveById(qgroup.Id);
                            this.ClientContext.ExecuteQueryRetry();
                        }
                        groupsFound--;
                    }

                    if(groupsFound == 0
                        && this.ShouldProcess(string.Format("Sharepoint user {0} will be removed from {1}", this.UserName, this.ClientContext.Url)))
                    {
                        this.ClientContext.Web.SiteUsers.RemoveById(quser.Id);
                        this.ClientContext.ExecuteQueryRetry();
                    }
                }
            }
            catch (Exception ex)
            {
                LogError(ex, "Failed to execute RemoveUserProfileFromSite by site {0}", this.ClientContext.Url);
            }
        }


    }
}
