using InfrastructureAsCode.Core.Extensions;
using InfrastructureAsCode.Powershell.Commands.Base;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Powershell.Commands.Sites
{
    /// <summary>
    /// The function cmdlet will enable or disable the modern UI
    /// </summary>
    /// <remarks>
    /// Turn off the Modern UI
    /// </remarks>
    [Cmdlet(VerbsCommon.Set, "IaCModernUI", SupportsShouldProcess = true)]
    public class SetIaCModernUI : IaCCmdlet
    {
        /// <summary>
        /// Should we enable modern UI
        /// </summary>
        [Parameter(Mandatory = false, HelpMessage = "If you pass this in, it will reenable ModernUI")]
        public SwitchParameter EnableModernUI { get; set; }

        /// <summary>
        /// Should we enable modern UI
        /// </summary>
        [Parameter(Mandatory = false, HelpMessage = "If you pass this in, it will assume its a web")]
        public SwitchParameter IsWeb { get; set; }


        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();

            var featureSiteGuid = new System.Guid("E3540C7D-6BEA-403C-A224-1A12EAFEE4C4");
            var featureWebGuid = new System.Guid("52E14B6F-B1BB-4969-B89B-C4FAA56745EF");

            // To apply the script to the site collection level, [set as default].
            var connectionUrl = this.ClientContext.Url;

            if (!EnableModernUI)
            {
                // To turn off the new UI by default in the new site, uncomment the next line.
                if (!IsWeb)
                {
                    this.ClientContext.Load(this.ClientContext.Site);
                    this.ClientContext.Site.ActivateFeature(featureSiteGuid);
                }
                else
                {
                    this.ClientContext.Load(this.ClientContext.Web);
                    this.ClientContext.Web.ActivateFeature(featureWebGuid);
                }
            }
            else
            {
                // To re-enable the option to use the new UI after having first disabled it, uncomment the next line.
                // and comment the preceding line.
                if (!IsWeb)
                {
                    this.ClientContext.Load(this.ClientContext.Site);
                    this.ClientContext.Site.DeactivateFeature(featureSiteGuid);
                }
                else
                {
                    this.ClientContext.Load(this.ClientContext.Web);
                    this.ClientContext.Web.DeactivateFeature(featureWebGuid);
                }
            }

            this.ClientContext.ExecuteQueryRetry();

        }
    }
}