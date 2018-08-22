using InfrastructureAsCode.Core.Models;
using InfrastructureAsCode.Powershell;
using InfrastructureAsCode.Powershell.Commands.Base;
using InfrastructureAsCode.Powershell.Extensions;
using Microsoft.Online.SharePoint.TenantAdministration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using PCommand = System.Management.Automation.Runspaces;

namespace InfrastructureAsCode.Powershell.Commands.Principals
{
    [Cmdlet(VerbsCommon.Get, "IaCPeopleManagerProfiles")]
    [CmdletHelp("Query for a user across the tenant", Category = "Principals")]
    public class GetIaCPeopleManagerProfiles : IaCCmdlet
    {

        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 1)]
        public string UserName { get; set; }


        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            if (string.IsNullOrEmpty(this.UserName))
            {
                throw new InvalidOperationException("UserName is not valid.");
            }
        }

        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();
            var models = new List<SPUserDefinitionModel>();

            try
            {
                LogVerbose("Querying UserProfiles.PeopleManager");
                var userPrincipalName = string.Format("i:0#.f|membership|{0}", this.UserName);
                var peopleContext = new Microsoft.SharePoint.Client.UserProfiles.PeopleManager(this.ClientContext);
                var personProperties = peopleContext.GetPropertiesFor(userPrincipalName);
                this.ClientContext.Load(personProperties);
                this.ClientContext.ExecuteQuery();

                if (personProperties != null)
                {
                    var profileProperties = personProperties.UserProfileProperties.ToList();
                    LogVerbose("Display Name: {0}", personProperties.DisplayName);
                    var model = new SPUserDefinitionModel()
                    {
                        UserDisplay = personProperties.DisplayName,
                        UserEmail = personProperties.Email,
                        LatestPost = personProperties.LatestPost,
                        OD4BUrl = personProperties.PersonalUrl,
                        UserProfileGUID = profileProperties.GetPropertyValue("UserProfile_GUID"),
                        SPSDistinguishedName = profileProperties.GetPropertyValue("SPS-DistinguishedName"),
                        SPSSid = profileProperties.GetPropertyValue("SID"),
                        MSOnlineObjectId = profileProperties.GetPropertyValue("msOnline-ObjectId")
                    };
                    models.Add(model);
                }

                models.ForEach(user => WriteObject(user));
            }
            catch (Exception ex)
            {
                LogError(ex, "Failed to execute QueryUserProfile in tenant {0}", this.ClientContext.Url);
            }
        }
    }
}
