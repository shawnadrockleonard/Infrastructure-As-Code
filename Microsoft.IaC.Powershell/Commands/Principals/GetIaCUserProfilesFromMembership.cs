using IaC.Powershell;
using IaC.Powershell.CmdLets;
using IaC.Core.Models;
using IaC.Core.Extensions;
using IaC.Powershell.Extensions;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using OfficeDevPnP.Core.Entities;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace IaC.Powershell.Commands.Principals
{
    /// <summary>
    /// This command will pull UPA information from a specific set of SharePoint Groups in a site
    /// </summary>
    [Cmdlet(VerbsCommon.Get, "IaCUserProfilesFromMembership")]
    [CmdletHelp("Retrieves UPA information based on the identified sharepoint group membership.", Category = "Principals")]
    public class GetIaCUserProfilesFromMembership : IaCCmdlet
    {
        /// <summary>
        /// Collection of SharePoint groups from which we will extract membership
        /// </summary>
        [Parameter(Mandatory = true, Position = 1)]
        public string[] SiteGroups { get; set; }

        /// <summary>
        /// Validate parameters
        /// </summary>
        protected override void BeginProcessing()
        {
            base.BeginProcessing();
        }

        /// <summary>
        /// Process the request
        /// </summary>
        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();

            if (this.ClientContext == null)
            {
                LogWarning("Invalid client context, configure the service to run again");
                return;
            }



            // obtain CSOM object for host web
            var vweb = this.ClientContext.Web;
            this.ClientContext.Load(vweb, hw => hw.SiteGroups, hw => hw.Title, hw => hw.ContentTypes);
            this.ClientContext.ExecuteQuery();


            GroupCollection groups = vweb.SiteGroups;
            this.ClientContext.Load(groups, g => g.Include(inc => inc.Id, inc => inc.Title, igg => igg.Users));
            this.ClientContext.ExecuteQuery();


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

            var siteGroupUsers = new List<SPUserDefinitionModel>();

            var filteredGroups = groups.Where(w => SiteGroups.Any(a => w.Title.Equals(a, StringComparison.CurrentCultureIgnoreCase)));

            foreach (var group in filteredGroups)
            {
                foreach (var user in group.Users)
                {
                    if (!siteGroupUsers.Any(a => a.UserName == user.LoginName))
                    {
                        var userProfile = ows.GetUserProfileByName(user.LoginName);
                        //var userOrgs = ows.GetUserOrganizations(user.LoginName);
                        var UserName = userProfile.RetrieveUserProperty("UserName");
                        var office = userProfile.RetrieveUserProperty("Department");
                        if (string.IsNullOrEmpty(office))
                            office = userProfile.RetrieveUserProperty("SPS-Department");

                        var userManager = userProfile.RetrieveUserProperty("Manager");

                        office = office.Replace(new char[] { '/', '\\', '-' }, ",");
                        office = office.Replace(" ", "");
                        var officeSplit = office.Split(new string[] { "," }, StringSplitOptions.None);
                        var officeAcronym = (officeSplit.Length > 0) ? officeSplit[0] : string.Empty;

                        siteGroupUsers.Add(new SPUserDefinitionModel()
                        {
                            Manager = userManager,
                            Organization = office,
                            OrganizationAcronym = officeAcronym,
                            UserName = user.LoginName,
                            UserEmail = user.Email,
                            UserDisplay = user.Title
                        });
                    }
                }



                WriteObject(siteGroupUsers);
            }

        }
    }
}
