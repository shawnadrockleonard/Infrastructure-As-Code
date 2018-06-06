using InfrastructureAsCode.Powershell;
using InfrastructureAsCode.Powershell.CmdLets;
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

            // Site Features
            var siteFeatures = ClientContext.LoadQuery(ClientContext.Site.Features.Include(fctx => fctx.DefinitionId, fctx => fctx.DisplayName));
            ClientContext.ExecuteQueryRetry();
            foreach (Feature SiteFeature in siteFeatures)
            {
                objects.Add(new FeatureDefinition()
                {
                    Id = SiteFeature.DefinitionId,
                    DisplayName = SiteFeature.DisplayName,
                    Scope = FeatureDefinitionScope.Site
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
                    Scope = FeatureDefinitionScope.Web
                });
            }

            // Export Objects
            WriteObject(objects, true);
        }

    }
}
