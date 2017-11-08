using InfrastructureAsCode.Core.Extensions;
using InfrastructureAsCode.Core.Models;
using InfrastructureAsCode.Powershell.CmdLets;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Powershell.Commands.Sites
{
    /// <summary>
    /// The function cmdlet will allow you to specify a JSON file and update the Site or Web with the appropriate User Custom Actions
    /// </summary>
    /// <remarks>
    /// Set-IaCCustomAction -FilePath c:\filedir\customaction.json -Verbose
    /// </remarks>
    [Cmdlet(VerbsCommon.Set, "IaCCustomAction", SupportsShouldProcess = true)]
    public class SetIaCCustomAction : IaCCmdlet
    {
        /// <summary>
        /// Full path to the JSON file
        /// </summary>
        [Parameter(Mandatory = false, ParameterSetName = "Default")]
        public string FilePath { get; set; }


        /// <summary>
        /// Evaluate the full path
        /// </summary>
        protected override void OnBeginInitialize()
        {
            if (this.ParameterSetName == "Default")
            {
                var fileInfo = new System.IO.FileInfo(FilePath);
                if (!fileInfo.Exists)
                {
                    throw new System.IO.FileNotFoundException("File not found", fileInfo.Name);
                }
            }
        }

        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();

            var site = this.ClientContext.Site;
            var web = this.ClientContext.Web;
            this.ClientContext.Load(site, ccsu => ccsu.ServerRelativeUrl, cssu => cssu.UserCustomActions);
            this.ClientContext.Load(web, ccwu => ccwu.ServerRelativeUrl, ccwu => ccwu.UserCustomActions);
            this.ClientContext.ExecuteQueryRetry();

            var siteurl = TokenHelper.EnsureTrailingSlash(site.ServerRelativeUrl);
            var weburl = TokenHelper.EnsureTrailingSlash(web.ServerRelativeUrl);


            var fileInfo = new System.IO.FileInfo(FilePath);
            var actions = JsonConvert.DeserializeObject<SPCustomAction>(System.IO.File.ReadAllText(fileInfo.FullName));
            if (actions.Site != null)
            {
                if (actions.Site.scriptblocks != null && actions.Site.scriptblocks.Any())
                {
                    actions.Site.scriptblocks.ForEach(cab =>
                    {
                        var htmlblock = cab.htmlblock.Replace("~SiteCollection/", siteurl);
                        htmlblock = htmlblock.Replace("~Site/", weburl);

                        site.AddOrUpdateCustomActionLinkBlock(cab.name, htmlblock, cab.sequence);
                    });
                }
                if (actions.Site.scriptlinks != null && actions.Site.scriptlinks.Any())
                {
                    actions.Site.scriptlinks.ForEach(cab =>
                    {
                        site.AddOrUpdateCustomActionLink(cab.name, cab.linkurl, cab.sequence);
                    });
                }
            }

            if (actions.Web != null)
            {
                if (actions.Web.scriptblocks != null && actions.Web.scriptblocks.Any())
                {
                    actions.Web.scriptblocks.ForEach(cab =>
                    {
                        var htmlblock = cab.htmlblock.Replace("~SiteCollection/", siteurl);
                        htmlblock = htmlblock.Replace("~Site/", weburl);

                        web.AddOrUpdateCustomActionLinkBlock(cab.name, htmlblock, cab.sequence);
                    });
                }
                if (actions.Web.scriptlinks != null && actions.Web.scriptlinks.Any())
                {
                    actions.Web.scriptlinks.ForEach(cab =>
                    {
                        web.AddOrUpdateCustomActionLink(cab.name, cab.linkurl, cab.sequence);
                    });
                }
            }

            if (actions.List != null && actions.List.Any())
            {
                foreach (var list in actions.List)
                {
                    var weblist = web.GetListByTitle(list.Title);
                    foreach (var listaction in list.scriptcommands)
                    {
                        var htmllink = listaction.ImageUrl.Replace("~SiteCollection/", siteurl);
                        htmllink = htmllink.Replace("~Site/", weburl);
                        listaction.ImageUrl = htmllink;
                        weblist.AddOrUpdateCustomActionLink(listaction);
                    }
                }
            }
        }
    }
}
