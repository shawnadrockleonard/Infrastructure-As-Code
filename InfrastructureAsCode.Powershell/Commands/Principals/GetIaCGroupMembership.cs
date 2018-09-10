using InfrastructureAsCode.Core.Models;
using InfrastructureAsCode.Powershell.Commands.Base;
using InfrastructureAsCode.Powershell.Extensions;
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using PCommand = System.Management.Automation.Runspaces;

namespace InfrastructureAsCode.Powershell.Commands.Principals
{
    [Cmdlet(VerbsCommon.Get, "IaCGroupMembership")]
    [CmdletHelp("Query membership for the group in the site", Category = "Principals")]
    public class GetIaCGroupMembership : IaCCmdlet
    {
        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 1)]
        public string GroupName { get; set; }

        [Parameter(Mandatory = false, ValueFromPipeline = true, Position = 2)]
        public SwitchParameter GroupMembership { get; set; }



        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            if (string.IsNullOrEmpty(this.GroupName))
            {
                throw new InvalidOperationException("GroupName is not valid.");
            }
        }

        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();
            var models = new List<SPGroupDefinitionModel>();

            try
            {
                var qroups = this.ClientContext.Web.SiteGroups;
                this.ClientContext.Load(qroups);

                var groupName = qroups.GetByName(this.GroupName);
                this.ClientContext.Load(groupName, gg => gg.LoginName, gg => gg.Title, gg => gg.Users);
                this.ClientContext.ExecuteQuery();

                var model = new SPGroupDefinitionModel()
                {
                    Title = groupName.LoginName,
                    Description = groupName.Title
                };

                LogVerbose("Group {0} found in Site {1}", groupName.Title, this.ClientContext.Web.Url);

                if (GroupMembership)
                {
                    foreach (var user in groupName.Users)
                    {
                        var userModel = new SPUserDefinitionModel()
                        {
                            UserName = user.LoginName,
                            UserEmail = user.Email,
                            GuidId = user.Id,
                            PrincipalType = user.PrincipalType,
                            UserDisplay = user.Title,
                            UserId = user.UserId
                        };
                        model.Users.Add(userModel);
                        LogVerbose("LoginName {0}", user.LoginName);
                    }
                }

                models.Add(model);

                models.ForEach(group => WriteObject(group));
            }
            catch (Exception ex)
            {
                LogError(ex, "Failed to execute QueryUserProfile in tenant {0}", this.ClientContext.Url);
            }
        }

    }
}
