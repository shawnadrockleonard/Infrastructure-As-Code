using InfrastructureAsCode.Powershell.CmdLets;
using InfrastructureAsCode.Core.HttpServices;
using InfrastructureAsCode.Core.Extensions;
using Microsoft.Online.SharePoint.TenantAdministration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using PCommand = System.Management.Automation.Runspaces;

namespace InfrastructureAsCode.Powershell.Commands.Principals
{
    /// <summary>
    /// Query for a user in the tenant
    /// </summary>
    [Cmdlet(VerbsCommon.Get, "IaCQueryProfileUser")]
    public class GetIaCQueryProfileUser : IaCCmdlet
    {

        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 1)]
        public string UserName { get; set; }

        [Parameter(Mandatory = false, ValueFromPipeline = true, Position = 2)]
        public string SiteUrl { get; set; }

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
            try
            {
                var userPrincipalName = string.Format("{0}|{1}", ClaimIdentifier, this.UserName);


                LogVerbose("Querying UserProfile service");
                GetUserProfileInformation(userPrincipalName);


                //load the tenant object
                var officeTenantContext = new Microsoft.Online.SharePoint.TenantManagement.Office365Tenant(this.ClientContext);
                var tenantContext = new Tenant(this.ClientContext);

                this.ClientContext.Load(officeTenantContext);
                this.ClientContext.Load(tenantContext);
                this.ClientContext.ExecuteQuery();

                LogVerbose("Querying UserProfiles.PeopleManager");
                GetUserInformation(userPrincipalName);


                using (var runspace = new SPIaCRunspaceWithDelegate(SPIaCConnection.CurrentConnection))
                {
                    LogVerbose("Executing runspace to query Get-MSOLUser(-UserPrincipalName {0})", this.UserName);

                    var getUserCommand = new PCommand.Command("Get-MsolUser");
                    getUserCommand.Parameters.Add((new PCommand.CommandParameter("UserPrincipalName", this.UserName)));

                    var results = runspace.ExecuteRunspace(getUserCommand, string.Format("Unable to get user with UserPrincipalName : " + userPrincipalName));
                    if (results.Count() > 0)
                    {
                        foreach (PSObject itemUser in results)
                        {
                            var userProperties = itemUser.Properties;

                            GetPSObjectValue(userProperties, "DisplayName");
                            GetPSObjectValue(userProperties, "FirstName");
                            GetPSObjectValue(userProperties, "LastName");
                            GetPSObjectValue(userProperties, "UserPrincipalName");
                            GetPSObjectValue(userProperties, "Department");
                            GetPSObjectValue(userProperties, "Country");
                            GetPSObjectValue(userProperties, "UsageLocation");
                        }
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

        private object GetUserInformation(string userPrincipalName)
        {
            var peopleContext = new Microsoft.SharePoint.Client.UserProfiles.PeopleManager(ClientContext);
            var personProperties = peopleContext.GetPropertiesFor(userPrincipalName);
            ClientContext.Load(personProperties);
            ClientContext.ExecuteQuery();
            if (personProperties != null)
            {
                LogVerbose("Tenant PeopleManager");
                var profileProperties = personProperties.UserProfileProperties.ToList();
                LogVerbose("Display Name: {0}", personProperties.DisplayName);
                LogVerbose("Email {0}", personProperties.Email);
                LogVerbose("Latest Post {0}", personProperties.LatestPost);
                LogVerbose("Personal Url {0}", personProperties.PersonalUrl);
                LogVerbose("UserProfile_GUID: {0}", GetPropertyValue(profileProperties, "UserProfile_GUID"));
                LogVerbose("SPS-DistinguisedName: {0}", GetPropertyValue(profileProperties, "SPS-DistinguishedName"));
                LogVerbose("SID: {0}", GetPropertyValue(profileProperties, "SID"));
                LogVerbose("msOnline-ObjectId: {0}", GetPropertyValue(profileProperties, "msOnline-ObjectId"));

                profileProperties.ForEach(f =>
                {
                    LogVerbose("Property:{0} with value:{1}", f.Key, f.Value);
                });
            }

            return null;
        }

        private object GetUserProfileInformation(string userPrincipalName)
        {
            using (var ups = new UserProfileService(this.ClientContext, this.ClientContext.Url))
            {
                var personProperties = ups.ows.GetUserProfileByName(userPrincipalName);
                if (personProperties != null)
                {
                    LogVerbose("User Profile Web Service");
                    LogVerbose("UserName: {0}", personProperties.RetrieveUserProperty("UserName"));
                    LogVerbose("FirstName: {0}", personProperties.RetrieveUserProperty("FirstName"));
                    LogVerbose("LastName: {0}", personProperties.RetrieveUserProperty("LastName"));
                    LogVerbose("PreferredName: {0}", personProperties.RetrieveUserProperty("PreferredName"));
                    LogVerbose("Manager: {0}", personProperties.RetrieveUserProperty("Manager"));
                    LogVerbose("Department: {0}", personProperties.RetrieveUserProperty("Department")); 
                    LogVerbose("SPS-Department: {0}", personProperties.RetrieveUserProperty("SPS-Department")); 
                    LogVerbose("WorkPhone {0}", personProperties.RetrieveUserProperty("WorkPhone"));
                    LogVerbose("OfficeCubicalLocation {0}", personProperties.RetrieveUserProperty("OfficeCubicalLocation"));
                    LogVerbose("BuldingLocation {0}", personProperties.RetrieveUserProperty("BuldingLocation"));

                    foreach (var userProp in personProperties)
                    {
                        LogVerbose("Property:{0} with value:{1}", userProp.Name, personProperties.RetrieveUserProperty(userProp.Name));
                    }
                }
            }

            return null;
        }
    }
}
