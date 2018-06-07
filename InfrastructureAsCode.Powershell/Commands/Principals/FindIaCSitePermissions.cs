using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Powershell.Commands.Principals
{
    using InfrastructureAsCode.Powershell;
    using InfrastructureAsCode.Powershell.CmdLets;
    using InfrastructureAsCode.Core.Models;
    using InfrastructureAsCode.Core.Extensions;
    using Microsoft.Online.SharePoint.TenantAdministration;
    using Microsoft.Online.SharePoint.TenantManagement;
    using Microsoft.SharePoint.Client;
    using OfficeDevPnP.Core.Utilities;

    /// <summary>
    /// This function will check a principals access to the specified site URL
    /// </summary>
    /// <remarks>Identify SharePoint groups that have the <see cref="UserName"/> added to</remarks>
    [Cmdlet(VerbsCommon.Find, "IaCSitePermissions", SupportsShouldProcess = true)]
    public class FindIaCSitePermissions : IaCAdminCmdlet
    {
        /// <summary>
        /// The site url to start checking
        /// </summary>
        [Parameter(Mandatory = false)]
        public string SiteUrl { get; set; }

        /// <summary>
        /// The principal identifier
        /// </summary>
        [Parameter(Mandatory = false)]
        public string UserName { get; set; }



        /// <summary>
        /// Collection of Known everyone groups
        /// </summary>
        private List<string> DiscoveryGroups { get; set; }


        private List<SPSiteModel> Model { get; set; }


        protected override void OnBeginInitialize()
        {
            base.OnBeginInitialize();


            DiscoveryGroups = new List<string>();
            Model = new List<SPSiteModel>();
        }

        public override void ExecuteCmdlet()
        {

            base.ExecuteCmdlet();
            try
            {
                TenantContext.EnsureProperties(tssp => tssp.RootSiteUrl);
                var TenantUrl = TenantContext.RootSiteUrl.EnsureTrailingSlashLowered();


                // Set the Auth Realm for the Tenant Web Context
                using (var siteWeb = this.ClientContext.Clone(SiteUrl))
                {
                    var cq = new ChangeQuery
                    {
                        FetchLimit = 10,
                        Item = true
                    };
                    var webAuthRealm = siteWeb.Web.GetChanges(cq);
                    siteWeb.Load(webAuthRealm);
                    siteWeb.ExecuteQuery();
                }

                var userPrincipalName = string.Format("{0}|{1}", ClaimIdentifier, this.UserName);
                DiscoveryGroups.Add(userPrincipalName);



                try
                {
                    SetSiteAdmin(SiteUrl, CurrentUserName, true);

                    using (var siteContext = this.ClientContext.Clone(SiteUrl))
                    {
                        Web _web = siteContext.Web;

                        ProcessSiteCollectionSubWeb(_web, true);

                        var siteProperties = GetSiteProperties(SiteUrl);
                    }
                }
                catch (Exception e)
                {
                    LogError(e, "Failed to processSiteCollection with url {0}", SiteUrl);
                }
                finally
                {
                    //SetSiteAdmin(_siteUrl, CurrentUserName);
                }

            }
            catch (Exception e)
            {
                LogError(e, "Failed in SetEveryoneGroup cmdlet {0}", e.Message);
            }

            WriteObject(Model);
        }

        private SiteProperties GetSiteProperties(string siteUrl)
        {
            var properties = TenantContext.GetSitePropertiesByUrl(siteUrl, true);
            TenantContext.Context.Load(properties);
            TenantContext.Context.ExecuteQueryRetry();
            return properties;
        }

        /// <summary>
        /// Process the site subweb
        /// </summary>
        /// <param name="_web"></param>
        /// <param name="isTopLevel">(OPTIONAL) indicates the top level</param>
        private ExtendSPOSiteModel ProcessSiteCollectionSubWeb(Web _web, bool isTopLevel = false)
        {
            try
            {
                _web.EnsureProperties(spp => spp.Id, spp => spp.Url);
                var _siteUrl = _web.Url;

                ProcessSite(_web);

                //Process subsites
                _web.EnsureProperties(spp => spp.Webs);

                if (_web.Webs.Count() > 1)
                {
                    LogVerbose("Site {0} has webs {1}", _siteUrl, _web.Webs.Count);
                    foreach (Web _inWeb in _web.Webs)
                    {
                        var model = ProcessSiteCollectionSubWeb(_inWeb, false);
                    }
                }
            }
            catch (Exception e)
            {
                LogError(e, "Failed in processSiteCollection");
            }

            return null;
        }


        private void ProcessSite(Web _web)
        {
            var site = new SPSiteModel();

            var admins = _web.GetAdministrators();


            _web.EnsureProperties(lssp => lssp.Id, wspp => wspp.ServerRelativeUrl, 
                wspp => wspp.Title, 
                spp => spp.HasUniqueRoleAssignments, 
                spp => spp.Url, 
                spp => spp.Lists);
            site.Url = _web.Url;
            site.title = _web.Title;

            LogVerbose("Processing: {0}", _web.Url);

            /* Process Site Owner */
            try
            {
                admins.ForEach(owner =>
                {
                    site.Owners.Add(new SPPrincipalModel()
                    {
                        LoginName = owner.LoginName
                    });
                });
            }
            catch (Exception e)
            {
                LogError(e, "Failed to retrieve site owners {0}", _web.Url);
            }

            // ********** Process Site user
            try
            {
                _web.EnsureProperties(wspp => wspp.RoleAssignments,
                spp => spp.SiteUsers.Include(sppi => sppi.Id, sppi => sppi.Title, sppi => sppi.LoginName));

                foreach (User _user in _web.SiteUsers)
                {
                    if (IsEveryoneInPrincipal(_user))
                    {
                        if (!site.Users.Any(u => u.Id == _user.Id))
                        {
                            var roles = _web.RoleAssignments.Where(ra => ra.PrincipalId == _user.Id).ToList();

                            _user.EnsureProperties(
                                susp => susp.Id,
                                susp => susp.Title,
                                susp => susp.LoginName,
                                susp => susp.IsHiddenInUI,
                                susp => susp.PrincipalType,
                                susp => susp.IsSiteAdmin,
                                susp => susp.Groups);

                            site.Users.Add(new SPPrincipalModel()
                            {
                                Id = _user.Id,
                                IsHiddenInUI = _user.IsHiddenInUI,
                                Title = _user.Title,
                                LoginName = _user.LoginName,
                                PrincipalType = _user.PrincipalType
                            });
                        }
                        break;
                    }
                }
            }
            catch (Exception e)
            {
                LogError(e, "Failed processUsers enumerating site users");
            }

            // ********** Process Site groups
            try
            {
                GroupCollection _groups = _web.SiteGroups;
                _web.Context.Load(_groups, gspp => gspp.Include(ussp => ussp.Id, ussp => ussp.Title, ussp => ussp.LoginName, ussp => ussp.IsHiddenInUI, ussp => ussp.Users));
                _web.Context.ExecuteQueryRetry();

                foreach (Group _group in _groups)
                {
                    UserCollection _users = _group.Users;
                    _group.EnsureProperties(usp => usp.Users);

                    foreach (User _xUser in _users)
                    {
                        if (IsEveryoneInPrincipal(_xUser))
                        {
                            if (!site.Groups.Any(g => g.Id == _xUser.Id))
                            {
                                site.Groups.Add(new SPGroupPrincipalModel()
                                {
                                    GroupId = _group.Id,
                                    GroupName = _group.Title,
                                    GroupHidden = _group.IsHiddenInUI,
                                    GroupLogin = _group.LoginName,
                                    Id = _xUser.Id,
                                    IsHiddenInUI = _xUser.IsHiddenInUI,
                                    Title = _xUser.Title,
                                    LoginName = _xUser.LoginName,
                                    PrincipalType = _xUser.PrincipalType
                                });
                            }
                            break;
                        }
                    }
                }
            }
            catch (Exception e)
            {
                LogError(e, "Failed in process Groups");
            }

            Model.Add(site);
        }


        /// <summary>
        /// Queries the collection of groups
        /// </summary>
        /// <param name="_principal"></param>
        /// <returns></returns>
        private bool IsEveryoneInPrincipal(Principal _principal)
        {
            if (DiscoveryGroups.Any(eg =>
                    _principal.Title.Equals(eg, StringComparison.CurrentCultureIgnoreCase)
                    || _principal.LoginName.Equals(eg, StringComparison.CurrentCultureIgnoreCase)))
            {
                return true;
            }
            return false;
        }


    }
}
