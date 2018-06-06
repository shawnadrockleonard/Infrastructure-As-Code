using InfrastructureAsCode.Powershell.CmdLets;
using InfrastructureAsCode.Core.Utilities;
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using PCommand = System.Management.Automation.Runspaces;

namespace InfrastructureAsCode.Powershell.Commands.Principals
{
    /// <summary>
    /// Query for a group in the tenant
    /// </summary>
    [Cmdlet(VerbsCommon.Get, "IaCQueryProfileGroup")]
    public class GetIaCQueryProfileGroup : IaCAdminCmdlet
    {
        [Parameter(Mandatory = false, ValueFromPipeline = true, Position = 1)]
        public string GroupId { get; set; }

        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 2)]
        public string GroupName { get; set; }

        [Parameter(Mandatory = false, ValueFromPipeline = true, Position = 3)]
        public SwitchParameter GroupMembership { get; set; }

        [Parameter(Mandatory = false, ValueFromPipeline = true, Position = 4)]
        public string SiteUrl { get; set; }

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
            try
            {

                using (var runspace = new SPIaCRunspaceWithDelegate(SPIaCConnection.CurrentConnection))
                {
                    if (string.IsNullOrEmpty(this.GroupId))
                    {
                        var getGroupAllCommand = new PCommand.Command("Get-MSOLGroup");
                        getGroupAllCommand.Parameters.Add((new PCommand.CommandParameter("All")));

                        LogVerbose("BEGIN ---------------");
                        LogVerbose("Executing runspace to query Get-MSOLGroup(-All) which will take a minute or two.");
                        var groupAllResults = runspace.ExecuteRunspace(getGroupAllCommand, string.Format("Unable to retrieve {0} groups", "All"));
                        if (groupAllResults.Count() > 0)
                        {
                            LogVerbose("MSOL Groups found {0}", groupAllResults.Count());
                        }
                        LogVerbose("END ---------------");
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(this.SiteUrl))
                        {
                            using (var thisContext = this.ClientContext.Clone(this.SiteUrl))
                            {
                                thisContext.Credentials = this.ClientContext.Credentials;

                                var qroups = thisContext.Web.SiteGroups;
                                thisContext.Load(qroups);

                                var groupName = qroups.GetByName(this.GroupName);
                                thisContext.Load(groupName, gg => gg.LoginName, gg => gg.Title, gg => gg.Users);
                                thisContext.ExecuteQuery();
                                LogVerbose("Group {0} found in Site {1}", groupName.Title, this.SiteUrl);
                                foreach(var user in groupName.Users)
                                {
                                    LogVerbose("LoginName {0}", user.LoginName);
                                    LogVerbose("Email {0}", user.Email);
                                    LogVerbose("Id {0}", user.Id);
                                    LogVerbose("PrincipalType {0}", user.PrincipalType);
                                    LogVerbose("Title {0}", user.Title);
                                    LogVerbose("UserId {0}", user.UserId);
                                }
                            }
                        }

                        var getGroupCommand = new PCommand.Command("Get-MSOLGroup");
                        getGroupCommand.Parameters.Add((new PCommand.CommandParameter("ObjectId", this.GroupId)));
                        getGroupCommand.Parameters.Add((new PCommand.CommandParameter("Verbose")));

                        LogVerbose("BEGIN ---------------");
                        LogVerbose("Executing runspace to query Get-MSOLGroup(-SearchString {0}).", this.GroupId);
                        var groupresults = runspace.ExecuteRunspace(getGroupCommand, string.Format("Unable to retrieve {0} groups", this.GroupId));
                        if (groupresults.Count() > 0)
                        {
                            foreach (var itemGroup in groupresults)
                            {
                                var groupProperties = itemGroup.Properties;
                                var groupObjectId = GetPSObjectValue(groupProperties, "ObjectId");

                                GetPSObjectValue(groupProperties, "CommonName");
                                GetPSObjectValue(groupProperties, "DisplayName");
                                GetPSObjectValue(groupProperties, "EmailAddress");
                                GetPSObjectValue(groupProperties, "GroupType");
                                GetPSObjectValue(groupProperties, "LastDirSyncTime");
                                GetPSObjectValue(groupProperties, "ManagedBy");
                                GetPSObjectValue(groupProperties, "ValidationStatus");
                                GetPSObjectValue(groupProperties, "IsSystem");

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

                                            GetPSObjectValue(memberProperties, "CommonName");
                                            GetPSObjectValue(memberProperties, "DisplayName");
                                            GetPSObjectValue(memberProperties, "EmailAddress");
                                            GetPSObjectValue(memberProperties, "GroupMemberType");
                                            GetPSObjectValue(memberProperties, "ObjectId");
                                        }
                                    }
                                    LogVerbose("END ---------------");
                                }
                            }
                        }
                        LogVerbose("END ---------------");
                    }
                }
            }
            catch (Exception ex)
            {
                LogError(ex, "Failed to execute QueryUserProfile in tenant {0}", this.ClientContext.Url);
            }
        }

        private string GetPSObjectValue(PSMemberInfoCollection<PSPropertyInfo> infoProperties, string propertyName)
        {
            var resultValue = string.Empty;
            if (infoProperties[propertyName] != null && infoProperties[propertyName].Value != null)
            {
                resultValue = infoProperties[propertyName].Value.ToString();
                LogVerbose("{0}: {1}", propertyName, resultValue);
            }
            return resultValue;
        }

        private static string GetPropertyValue(List<KeyValuePair<string, string>> profileProperties, string propertyName, Type valueType = null)
        {
            var property = profileProperties.FirstOrDefault(f => f.Key.Equals(propertyName, StringComparison.CurrentCultureIgnoreCase));
            if (!property.Equals(default(KeyValuePair<string, string>)))
            {
                if (valueType != null && valueType == typeof(System.Guid))
                {
                    var propValue = new Guid(property.Value);
                    return propValue.ToString("D");
                }
                return property.Value;
            }
            return string.Empty;
        }
    }
}
