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
    [Cmdlet(VerbsCommon.Get, "IaCMSOnlineUserProfiles")]
    [CmdletHelp("Query for a user in the tenant", Category = "Principals")]
    public class GetIaCMSOnlineUserProfiles : IaCCmdlet
    {

        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 1)]
        public string UserName { get; set; }



        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();
            try
            {
                using (var runspace = new SPIaCRunspaceWithDelegate())
                {
                    runspace.Initialize(SPIaCConnection.CurrentConnection, "MSOnline", "Connect-MsolService");


                    LogVerbose("Executing runspace to query Get-MSOLUser(-UserPrincipalName {0})", this.UserName);

                    var getUserCommand = new PCommand.Command("Get-MsolUser");
                    getUserCommand.Parameters.Add((new PCommand.CommandParameter("UserPrincipalName", this.UserName)));

                    var results = runspace.ExecuteRunspace(getUserCommand, string.Format("Unable to get user with UserPrincipalName : {0}", UserName));
                    if (results.Count() > 0)
                    {
                        foreach (PSObject itemUser in results)
                        {
                            var userProperties = itemUser.Properties;

                            userProperties.GetPSObjectValue("DisplayName");
                            userProperties.GetPSObjectValue("FirstName");
                            userProperties.GetPSObjectValue("LastName");
                            userProperties.GetPSObjectValue("UserPrincipalName");
                            userProperties.GetPSObjectValue("Department");
                            userProperties.GetPSObjectValue("Country");
                            userProperties.GetPSObjectValue("UsageLocation");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogError(ex, "Failed to execute QueryUserProfile in tenant {0}", this.ClientContext.Url);
            }
        }
    }
}
