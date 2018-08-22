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
    /// The function cmdlet will set the laternate CSS for the site
    /// </summary>
    /// <remarks>
    ///     Set-IaCAlternateCSS -FileServerRelativeUrl "/SiteAssets/CSS/NewSite.css" will set the alternate CSS to run on every page
    /// </remarks>
    [Cmdlet(VerbsCommon.Set, "IaCAlternateCSS", SupportsShouldProcess = true)]
    public class SetIaCAlternateCSS : IaCCmdlet
    {
        /// <summary>
        /// Specific view to be updated from the above action list
        /// </summary>
        [Parameter(Mandatory = true)]
        public string FileServerRelativeUrl { get; set; }

        /// <summary>
        /// Validate parameters
        /// </summary>
        /// 
        protected override void OnBeginInitialize()
        {
            if (string.IsNullOrEmpty(FileServerRelativeUrl))
            {
                LogWarning("Failed to set alternate css url as one of the specified parameters was empty SpecificListName or SpecificFileName");
                return;
            }
        }

        /// <summary>
        /// Process the request
        /// </summary>
        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();


            var clientWeb = ClientContext.Web;
            if (!clientWeb.IsPropertyAvailable(wde => wde.AlternateCssUrl) || !ClientContext.Web.IsPropertyAvailable(wde => wde.UserCustomActions))
            {
                ClientContext.Load(clientWeb, w => w.AlternateCssUrl, w => w.UserCustomActions);
                ClientContext.ExecuteQueryRetry();
            }

            LogVerbose("Previous alternate CSS Url is {0}", clientWeb.AlternateCssUrl);

            FileServerRelativeUrl = (FileServerRelativeUrl.StartsWith("/") ? FileServerRelativeUrl.Substring(1) : FileServerRelativeUrl);
            var webUrl = new Uri(this.ClientContext.Url);
            var fileUrl = new Uri(webUrl, FileServerRelativeUrl);

            var baseUrl = fileUrl.GetLeftPart(UriPartial.Authority);
            var fileServerRelativeUrl = fileUrl.ToString().Replace(baseUrl, string.Empty);
            var file = clientWeb.GetFileByServerRelativeUrl(fileServerRelativeUrl);
            ClientContext.Load(file);
            ClientContext.ExecuteQueryRetry();


            if (ShouldProcess(string.Format("Processing alternate CSS URL {0}", fileServerRelativeUrl)))
            {
                clientWeb.AlternateCssUrl = file.ServerRelativeUrl;
                clientWeb.Update();
                ClientContext.ExecuteQueryRetry();
            }
        }
    }
}