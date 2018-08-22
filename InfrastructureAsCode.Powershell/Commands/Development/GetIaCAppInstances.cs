using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Powershell.Commands.Development
{
    using InfrastructureAsCode.Powershell.Commands.Base;
    using Microsoft.SharePoint.Client;

    /// <summary>
    /// Opens a web and queries apps
    /// </summary>
    [Cmdlet(VerbsCommon.Get, "IaCAppInstances")]
    public class GetIaCAppInstances : IaCCmdlet
    {
        [Parameter(Mandatory = false)]
        public Guid AppInstanceId { get; set; }

        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();
            try
            {
                
                var ctx = this.ClientContext;
                var apps = AppCatalog.GetAppInstances(ctx, this.ClientContext.Web);
                var result = ctx.LoadQuery(apps.Where(a => a.Status == AppInstanceStatus.Installed));
                this.ClientContext.ExecuteQueryRetry();
                foreach(AppInstance res in result)
                {
                    LogVerbose($"AppCatalog => App {res.Id} instance id {res.Title} and catalog {res.RemoteAppUrl} w/ SvcPri {res.AppPrincipalId}");
                }


                if (AppInstanceId != Guid.Empty)
                {
                    var appPermissions = ctx.Web.GetAppInstanceById(AppInstanceId);
                    ctx.Load(appPermissions);
                    ctx.ExecuteQuery();

                   var appd =  AppCatalog.GetAppPermissionDescriptions(ctx, ctx.Web, appPermissions);
                    this.ClientContext.ExecuteQueryRetry();
                }
            }
            catch (Exception ex)
            {
                LogError(ex, $"Failed in EPAAppInstances for {ex.Message}");
    }
}
    }
}
