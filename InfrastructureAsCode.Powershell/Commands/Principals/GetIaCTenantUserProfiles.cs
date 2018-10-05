using InfrastructureAsCode.Core;
using InfrastructureAsCode.Core.Extensions;
using InfrastructureAsCode.Core.HttpServices;
using InfrastructureAsCode.Core.Models;
using InfrastructureAsCode.Powershell.Commands.Base;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
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
                TenantContext.EnsureProperties(tssp => tssp.RootSiteUrl);
                var TenantUrl = TenantContext.RootSiteUrl.EnsureTrailingSlashLowered();
                var MySiteTenantUrl = TenantUrl.Replace(".sharepoint.com", "-my.sharepoint.com");

                var i = 1;
                var ilogger = new DefaultUsageLogger(LogVerbose, LogWarning, LogError);
                var odfb = this.ClientContext.GetOneDriveSiteCollections(ilogger, MySiteTenantUrl, true);

                models = odfb.Select(s =>
                {
                    var UserName = s.UserName;
                    if (!string.IsNullOrEmpty(UserName))
                    {
                        UserName = UserName.ToString().Replace(";", ",");
                    }

                    var userObject = new SPUserDefinitionModel()
                    {
                        UserName = UserName,
                        OD4BUrl = s.Url,
                        UserIndex = i++
                    };

                    return userObject;

                }).ToList();

            }
            catch (Exception ex)
            {
                LogError(ex, "Failed to retreive user profiles");
            }

            WriteObject(models, true);
        }

    }
}
