using InfrastructureAsCode.Powershell;
using InfrastructureAsCode.Powershell.Commands.Base;
using InfrastructureAsCode.Core.Extensions;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;
using InfrastructureAsCode.Core.Models;

namespace InfrastructureAsCode.Powershell.Commands.Features
{
    /// <summary>
    /// Reports the features
    /// </summary>
    [Cmdlet(VerbsCommon.Get, "IaCFeatures")]
    public class GetIaCFeatures : IaCCmdlet
    {

        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();

            var objects = new List<FeatureDefinition>();

            var web = this.ClientContext.Web;
            var site = this.ClientContext.Site;
            this.ClientContext.Load(web);
            this.ClientContext.Load(site);
            this.ClientContext.ExecuteQueryRetry();


            // Site Features
            var siteFeatures = ClientContext.LoadQuery(ClientContext.Site.Features.Include(fctx => fctx.DefinitionId, fctx => fctx.DisplayName));
            ClientContext.ExecuteQueryRetry();
            foreach (Feature siteFeature in siteFeatures)
            {
                objects.Add(new FeatureDefinition()
                {
                    Id = siteFeature.DefinitionId,
                    DisplayName = siteFeature.DisplayName,
                    Scope = FeatureDefinitionScope.Site,
                    IsActivated = FeatureExtensions.IsFeatureActive(site, siteFeature.DefinitionId)
                });
            }

            // Web Features
            var webFeatures = ClientContext.LoadQuery(ClientContext.Web.Features.Include(fctx => fctx.DefinitionId, fctx => fctx.DisplayName));
            ClientContext.ExecuteQueryRetry();
            foreach (var webFeature in webFeatures)
            {
                objects.Add(new FeatureDefinition()
                {
                    Id = webFeature.DefinitionId,
                    DisplayName = webFeature.DisplayName,
                    Scope = FeatureDefinitionScope.Web,
                    IsActivated = FeatureExtensions.IsFeatureActive(web, webFeature.DefinitionId)
                });
            }

            // Export Objects
            WriteObject(objects, true);
        }

    }
}
