using InfrastructureAsCode.Core.Extensions;
using InfrastructureAsCode.Core.HttpServices;
using InfrastructureAsCode.Core.Models;
using InfrastructureAsCode.Powershell.Commands.Base;
using System;
using System.Collections.Generic;
using System.Management.Automation;

namespace InfrastructureAsCode.Powershell.Commands.Principals
{
    [Cmdlet(VerbsCommon.Get, "IaCTenantUserProfiles")]
    [CmdletHelp("Opens a administrative web request and queries the user profile service", Category = "Principals")]
    public class GetIaCTenantUserProfiles : IaCAdminCmdlet
    {

        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();
            var models = new List<SPUserDefinitionModel>();

            try
            {
                var userProfile = new UserProfileService(this.ClientContext);

                var UserProfileResult = userProfile.OWService.GetUserProfileByIndex(-1);
                var NumProfiles = userProfile.OWService.GetUserProfileCount();

                var i = 1;
                var tmpCount = 0;
                var nextValue = UserProfileResult.NextValue;
                var nextValueIndex = int.Parse(nextValue);

                // As long as the next User profile is NOT the one we started with (at -1)...
                while (nextValueIndex != -1)
                {
                    LogVerbose("Examining profile {0} of {1}", i, NumProfiles);

                    // Look for the Personal Space object in the User Profile and retrieve it
                    // (PersonalSpace is the name of the path to a user's OneDrive for Business site. Users who have not yet created a 
                    // OneDrive for Business site might not have this property set.)

                    tmpCount++;

                    var PersonalSpaceUrl = UserProfileResult.RetrieveUserProperty("PersonalSpace");
                    var UserName = UserProfileResult.RetrieveUserProperty("UserName");
                    if (!string.IsNullOrEmpty(UserName))
                    {
                        UserName = UserName.ToString().Replace(";", ",");
                    }
                    
                    var userObject = new SPUserDefinitionModel()
                    {
                        UserName = UserName,
                        OD4BUrl = PersonalSpaceUrl,
                        UserIndex = nextValueIndex
                    };
                    models.Add(userObject);

                    // And now we check the next profile the same way...
                    UserProfileResult = userProfile.OWService.GetUserProfileByIndex(nextValueIndex);
                    nextValue = UserProfileResult.NextValue;
                    nextValueIndex = int.Parse(nextValue);
                    i++;
                }


                WriteObject(models, true);
            }
            catch (Exception ex)
            {
                LogError(ex, "Failed to retreive user profiles");
            }
        }

    }
}
