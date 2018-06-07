using InfrastructureAsCode.Powershell;
using InfrastructureAsCode.Powershell.CmdLets;
using InfrastructureAsCode.Core.Extensions;
using InfrastructureAsCode.Core.HttpServices;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Powershell.Commands.Principals
{
    /// <summary>
    /// Opens a administrative web request and queries the user profile service
    /// </summary>
    [Cmdlet(VerbsCommon.Get, "IaCQueryUserProfile")]
    public class GeIaCQueryUserProfile : IaCAdminCmdlet
    {

        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();
            try
            {


                TenantContext.EnsureProperties(tctx => tctx.RootSiteUrl);
                var rootMySiteUrl = TenantContext.RootSiteUrl.Replace(".sharepoint.com", "-my.sharepoint.com");


                using (var ups = new UserProfileService(this.ClientContext, this.ClientContext.Url))
                {
                    
                    var UserProfileResult = ups.OWService.GetUserProfileByIndex(-1);
                    var NumProfiles = ups.OWService.GetUserProfileCount();
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

                        var odbFound = false;
                        tmpCount++;

                        var Url = UserProfileResult.RetrieveUserProperty("PersonalSpace");
                        if (!string.IsNullOrEmpty(Url))
                        {
                            Url = rootMySiteUrl + Url;
                            odbFound = true;
                        }

                        var UserName = UserProfileResult.RetrieveUserProperty("UserName");
                        if (!string.IsNullOrEmpty(UserName))
                        {
                            UserName = UserName.ToString().Replace(";", ",");
                        }

                        LogVerbose("MainLine {0} URL:{1} OD4B Found:{2} UserName:{3}", nextValueIndex, Url, odbFound, UserName);

                        var userObject = new
                        {
                            UserName = UserName,
                            OD4BUrl = Url,
                            UserIndex = nextValueIndex
                        };
                        WriteObject(userObject);

                        // And now we check the next profile the same way...
                        UserProfileResult = ups.OWService.GetUserProfileByIndex(nextValueIndex);
                        nextValue = UserProfileResult.NextValue;
                        nextValueIndex = int.Parse(nextValue);
                        i++;
                    }

                }
            }
            catch (Exception ex)
            {
                LogError(ex, "Failed to retreive user profiles");
            }
        }

    }
}
