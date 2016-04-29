using IaC.Core.Models;
using IaC.Powershell.CmdLets;
using IaC.Powershell.Extensions;
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Management.Automation;
using PCommand = System.Management.Automation.Runspaces;

namespace IaC.Powershell.Commands.Runspace
{
    [Cmdlet(VerbsCommon.Get, "IaCMSOnlineProfileGroup")]
    [CmdletHelp("Query for a group in the tenant", Category = "Runspace")]
    public class GetIaCMSOnlineGroupMembership : IaCCmdlet
    {
        [Parameter(Mandatory = false, ValueFromPipeline = true, Position = 1)]
        public string GroupId { get; set; }

        [Parameter(Mandatory = false, ValueFromPipeline = true, Position = 2)]
        public SwitchParameter GroupMembership { get; set; }


        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();
            var models = new List<SPGroupDefinitionModel>();

            try
            {
                using (var runspace = new SPIaCRunspaceWithDelegate())
                {
                    runspace.Initialize(SPIaCConnection.CurrentConnection, "MSOnline", "Connect-MsolService");

                    var getGroupCommand = new PCommand.Command("Get-MSOLGroup");

                    if (string.IsNullOrEmpty(this.GroupId))
                    {
                        getGroupCommand.Parameters.Add((new PCommand.CommandParameter("All")));
                    }
                    else
                    {
                        getGroupCommand.Parameters.Add((new PCommand.CommandParameter("ObjectId", this.GroupId)));
                        getGroupCommand.Parameters.Add((new PCommand.CommandParameter("Verbose")));
                    }


                    LogVerbose("BEGIN ---------------");
                    LogVerbose("Executing runspace to query Get-MSOLGroup(-All) which will take a minute or two.");
                    var collectionOfGroups = runspace.ExecuteRunspace(getGroupCommand, string.Format("Unable to retrieve {0} groups", "All"));
                    LogVerbose("END ---------------");

                    if (collectionOfGroups.Count() > 0)
                    {
                        LogVerbose("MSOL Groups found {0}", collectionOfGroups.Count());

                        foreach (var itemGroup in collectionOfGroups)
                        {
                            var groupProperties = itemGroup.Properties;
                            var groupObjectId = groupProperties.GetPSObjectValue("ObjectId");

                            var model = new SPGroupMSOnlineDefinition()
                            {
                                ObjectId = groupObjectId,
                                Title = groupProperties.GetPSObjectValue("CommonName"),
                                Description = groupProperties.GetPSObjectValue("DisplayName"),
                                EmailAddress = groupProperties.GetPSObjectValue("EmailAddress"),
                                GroupType = groupProperties.GetPSObjectValue("GroupType"),
                                LastDirSyncTime = groupProperties.GetPSObjectValue("LastDirSyncTime"),
                                ManagedBy = groupProperties.GetPSObjectValue("ManagedBy"),
                                ValidationStatus = groupProperties.GetPSObjectValue("ValidationStatus"),
                                IsSystem = groupProperties.GetPSObjectValue("IsSystem")
                            };


                            if (GroupMembership)
                            {
                                var getGroupMembershipCommand = new PCommand.Command("Get-MsolGroupMember");
                                getGroupMembershipCommand.Parameters.Add((new PCommand.CommandParameter("GroupObjectId", groupObjectId)));
                                getGroupMembershipCommand.Parameters.Add((new PCommand.CommandParameter("Verbose")));

                                LogVerbose("BEGIN ---------------");
                                LogVerbose("Executing runspace to query Get-MsolGroupMember(-GroupObjectId {0}).", groupObjectId);
                                var groupMembershipResults = runspace.ExecuteRunspace(getGroupMembershipCommand, string.Format("Unable to retrieve {0} group membership", this.GroupId));
                                if (groupMembershipResults.Count() > 0)
                                {
                                    foreach (var itemMember in groupMembershipResults)
                                    {
                                        var memberProperties = itemMember.Properties;
                                        var userModel = new SPUserDefinitionModel()
                                        {
                                            UserName = memberProperties.GetPSObjectValue("CommonName"),
                                            UserDisplay = memberProperties.GetPSObjectValue("DisplayName"),
                                            UserEmail = memberProperties.GetPSObjectValue("EmailAddress"),
                                            Organization = memberProperties.GetPSObjectValue("GroupMemberType"),
                                            MSOnlineObjectId = memberProperties.GetPSObjectValue("ObjectId")
                                        };
                                        model.Users.Add(userModel);
                                    }
                                }
                            }

                            models.Add(model);
                        }
                    }

                }

                models.ForEach(groups => WriteObject(groups));
            }
            catch (Exception ex)
            {
                LogError(ex, "Failed to execute QueryUserProfile in tenant {0}", this.ClientContext.Url);
            }
        }
    }
}
