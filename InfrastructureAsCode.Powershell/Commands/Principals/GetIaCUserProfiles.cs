using InfrastructureAsCode.Powershell;
using InfrastructureAsCode.Powershell.CmdLets;
using InfrastructureAsCode.Core.Extensions;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;
using InfrastructureAsCode.Core.Models;

namespace InfrastructureAsCode.Powershell.Commands.Principals
{
    [Cmdlet(VerbsCommon.Get, "IaCUserProfiles")]
    [CmdletHelp("Opens a administrative web request and queries the user profile service", Category = "Principals")]
    public class GetIaCUserProfiles : SPOAdminCmdlet
    {

        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();
            var models = new List<SPUserDefinitionModel>();

            try
            {
                var creds = SPIaCConnection.CurrentConnection.GetActiveCredentials();
                var newcreds = new System.Net.NetworkCredential(creds.UserName, creds.Password);
                var spourl = new Uri(this.ClientContext.Url);
                var spocreds = new Microsoft.SharePoint.Client.SharePointOnlineCredentials(creds.UserName, creds.Password);
                var spocookies = spocreds.GetAuthenticationCookie(spourl);

                var spocontainer = new System.Net.CookieContainer();
                spocontainer.SetCookies(spourl, spocookies);

                var ows = new OfficeDevPnP.Core.UPAWebService.UserProfileService();
                ows.Url = string.Format("{0}/_vti_bin/userprofileservice.asmx", spourl.AbsoluteUri);
                ows.Credentials = newcreds;
                ows.CookieContainer = spocontainer;
                var UserProfileResult = ows.GetUserProfileByIndex(-1);
                var NumProfiles = ows.GetUserProfileCount();
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
                    UserProfileResult = ows.GetUserProfileByIndex(nextValueIndex);
                    nextValue = UserProfileResult.NextValue;
                    nextValueIndex = int.Parse(nextValue);
                    i++;
                }

                models.ForEach(user => WriteObject(user));
            }
            catch (Exception ex)
            {
                LogError(ex, "Failed to retreive user profiles");
            }
        }

    }
}
