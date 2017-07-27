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
    /// The function cmdlet will retrieve site/web custom actions and write them to the specified JSON file
    /// </summary>
    [Cmdlet(VerbsCommon.Copy, "IaCCustomAction", SupportsShouldProcess = true)]
    public class CopyIaCCustomAction : IaCCmdlet
    {

        [Parameter(Mandatory = true)]
        public string FilePath { get; set; }

        protected override void OnBeginInitialize()
        {
            var fileInfo = new System.IO.FileInfo(FilePath);
            if (!fileInfo.Directory.Exists)
            {
                throw new System.IO.DirectoryNotFoundException(string.Format("Directory not found => {0}", fileInfo.Directory.FullName));
            }
        }

        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();

            var fileInfo = new System.IO.FileInfo(FilePath);
            var site = this.ClientContext.Site;
            var web = this.ClientContext.Web;
            this.ClientContext.Load(site, ccsu => ccsu.ServerRelativeUrl, cssu => cssu.UserCustomActions);
            this.ClientContext.Load(web, ccwu => ccwu.ServerRelativeUrl, ccwu => ccwu.UserCustomActions);
            this.ClientContext.ExecuteQueryRetry();

            var siteurl = TokenHelper.EnsureTrailingSlash(site.ServerRelativeUrl);
            var weburl = TokenHelper.EnsureTrailingSlash(web.ServerRelativeUrl);


            var actions = new SPCustomAction();
            if (site.UserCustomActions != null && site.UserCustomActions.Any())
            {
                actions.Site = new SPCustomActionScope();

                foreach (var customAction in site.UserCustomActions)
                {
                    if (!string.IsNullOrEmpty(customAction.ScriptBlock))
                    {
                        if (actions.Site.scriptblocks == null)
                        {
                            actions.Site.scriptblocks = new List<SPCustomActionBlock>();
                        }

                        actions.Site.scriptblocks.Add(new SPCustomActionBlock()
                        {
                            name = customAction.Name,
                            htmlblock = customAction.ScriptBlock.Replace(siteurl, "~SiteCollection/").Replace(weburl, "~Site/"),
                            sequence = customAction.Sequence
                        });
                    }

                    if (!string.IsNullOrEmpty(customAction.ScriptSrc))
                    {
                        if (actions.Site.scriptlinks == null)
                        {
                            actions.Site.scriptlinks = new List<SPCustomActionLink>();
                        }

                        actions.Site.scriptlinks.Add(new SPCustomActionLink()
                        {
                            name = customAction.Name,
                            linkurl = customAction.ScriptSrc,
                            sequence = customAction.Sequence
                        });
                    }
                }
            }

            if (web.UserCustomActions != null && web.UserCustomActions.Any())
            {
                actions.Web = new SPCustomActionScope();

                foreach (var customAction in site.UserCustomActions)
                {
                    if (!string.IsNullOrEmpty(customAction.ScriptBlock))
                    {
                        if (actions.Web.scriptblocks == null)
                        {
                            actions.Web.scriptblocks = new List<SPCustomActionBlock>();
                        }

                        actions.Web.scriptblocks.Add(new SPCustomActionBlock()
                        {
                            name = customAction.Name,
                            htmlblock = customAction.ScriptBlock.Replace(siteurl, "~SiteCollection/").Replace(weburl, "~Site/"),
                            sequence = customAction.Sequence
                        });
                    }

                    if (!string.IsNullOrEmpty(customAction.ScriptSrc))
                    {
                        if (actions.Web.scriptlinks == null)
                        {
                            actions.Web.scriptlinks = new List<SPCustomActionLink>();
                        }

                        actions.Web.scriptlinks.Add(new SPCustomActionLink()
                        {
                            name = customAction.Name,
                            linkurl = customAction.ScriptSrc,
                            sequence = customAction.Sequence
                        });
                    }
                }
            }

            // write the actions to disk
            WriteCustomActionJsonToDisk(fileInfo, actions);
            WriteObject(actions);
        }

        private void WriteCustomActionJsonToDisk(System.IO.FileInfo fileInfo, SPCustomAction actions)
        {
            var jsonsettings = new JsonSerializerSettings()
            {
                Formatting = Formatting.Indented,
                Culture = System.Globalization.CultureInfo.CurrentUICulture,
                DateFormatHandling = DateFormatHandling.IsoDateFormat,
                NullValueHandling = NullValueHandling.Ignore
            };
            var actionjson = JsonConvert.SerializeObject(actions, jsonsettings);
            System.IO.File.WriteAllText(fileInfo.FullName, actionjson);
        }

    }
}
