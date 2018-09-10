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
    [Cmdlet(VerbsCommon.Get, "IaCExternalUserFromSite")]
    [CmdletHelp("retreives external user from the sharepoint site and tenant collection", Category = "External")]
    public class GetIaCExternalUserFromSite : IaCAdminCmdlet
    {

        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
        public string UserName { get; set; }


        [Parameter(Mandatory = false, ValueFromPipeline = true, Position = 1)]
        public string SiteUrl { get; set; }



        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();

            var invitedAs = string.Empty;
            var externalUsers = new List<SPExternalUserEntity>();
            var ilogger = new DefaultUsageLogger(LogVerbose, LogWarning, LogError);


            try
            {
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
               

                if (!string.IsNullOrEmpty(this.SiteUrl))
                {
                    externalUsers = this.ClientContext.CheckExternalUserForSite(ilogger, this.SiteUrl, invitedAs);
                }

                if (!externalUsers.Any())
                {
                    externalUsers = this.ClientContext.CheckExternalUser(ilogger, invitedAs);
                }


                if (!string.IsNullOrEmpty(this.SiteUrl))
                {
                    using (var thisContext = this.ClientContext.Clone(this.SiteUrl))
                    {

                        var siteUsrs = thisContext.Web.SiteUsers;
                        thisContext.Load(siteUsrs);
                        thisContext.ExecuteQueryRetry();

                        foreach (var externalUsr in externalUsers)
                        {

                            var foundUser = siteUsrs.FirstOrDefault(u => u.LoginName == externalUsr.AcceptedAs
                            || u.LoginName == externalUsr.InvitedAs
                            || u.Email == externalUsr.InvitedAs
                            || u.Title == externalUsr.InvitedAs);
                            if (foundUser == null)
                            {
                                LogVerbose($"The user {invitedAs} could not be found in the Site Users collection.");
                            }
                            else
                            {
                                externalUsr.FoundInSiteUsers = true;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogError(ex, "Failed to execute RemoveUserProfileFromSite by site {0}", this.SiteUrl);
            }

            WriteObject(externalUsers);
        }


    }
}
