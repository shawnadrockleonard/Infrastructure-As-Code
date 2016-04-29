using IaC.Core.Models;
using IaC.Powershell.CmdLets;
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;

namespace IaC.Powershell.Commands.Reporting
{
    [Cmdlet(VerbsCommon.Get, "IaCAnalytics")]
    [CmdletHelp("Returns a report of the entire farm", Category = "Reporting")]
    public class GetIaCAnalytics : SPOAdminCmdlet
    {
        /// <summary>
        /// The absolute URL to the site collection or web
        /// </summary>
        [Parameter(Mandatory = true, HelpMessage = "Provides a specific site to query and manipulate")]
        public string SiteUrl { get; set; }

        public List<SPWebDefinitionModel> WebModels { get; set; }

        /// <summary>
        /// Will enumerate the farm and return diagnostics and auditing information
        /// </summary>
        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();
            this.WebModels = new List<SPWebDefinitionModel>();

            try
            {
                ProcessSite(SiteUrl);

                WebModels.ForEach(web => WriteObject(web));
            }
            catch (Exception e)
            {
                LogError(e, "Failed in execute cmdlet with script option {0}", this.SiteUrl);
            }
        }

        /// <summary>
        /// Capture Reporting Data starting at this site url
        /// </summary>
        /// <param name="_siteUrl"></param>
        internal virtual void ProcessSite(string _siteUrl)
        {
            try
            {
                WebCollection subWebs = null;

                using (var ctx = this.ClientContext.Clone(_siteUrl))
                {
                    Web _web = ctx.Web;
                    ctx.Load(_web);
                    ctx.ExecuteQuery();

                    Site _site = ctx.Site;
                    subWebs = _web.Webs;

                    ctx.Load(_web, s => s.UIVersion, s => s.LastItemModifiedDate, s => s.Created, s => s.AssociatedOwnerGroup, s => s.Lists.Include(i => i.Id, i => i.Title, i => i.BaseType, i => i.ItemCount));
                    ctx.Load(subWebs);
                    ctx.Load(_site, s => s.Usage);
                    ctx.ExecuteQueryRetry();

                    var siteOwner = string.Empty;

                    try
                    {
                        var _user = _site.Owner;
                        ctx.Load(_user, iu => iu.Email);
                        ctx.ExecuteQueryRetry();
                        siteOwner = _user.Email;
                    }
                    catch (Exception e)
                    {
                        LogError(e, "Failed to retrieve owner from site {0}", _siteUrl);
                    }

                    var totalListCount = _web.Lists.Count();
                    var totalListItemCount = _web.Lists.Sum(l => l.ItemCount);
                    UsageInfo _usageInfo = _site.Usage;


                    var model = new SPWebDefinitionModel()
                    {
                        SiteUrl = _siteUrl,
                        SiteOwner = siteOwner,
                        UIVersion = _web.UIVersion,
                        Created = _web.Created,
                        UsageInfo = _usageInfo,
                        LastItemModifiedDate = _web.LastItemModifiedDate,
                        ListCount = totalListCount,
                        ListItemCount = totalListItemCount
                    };

                    WebModels.Add(model);
                }

                if (subWebs != null && subWebs.Count() > 0)
                {
                    foreach (var subWeb in subWebs)
                    {
                        ProcessSite(subWeb.Url);
                    }
                }
            }
            catch (Exception e)
            {
                LogError(e, "Failed in ProcessSite({0})", _siteUrl);
            }
        }
    }
}
