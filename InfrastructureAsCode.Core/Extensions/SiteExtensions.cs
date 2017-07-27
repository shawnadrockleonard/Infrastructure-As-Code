using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Extensions
{
    /// <summary>
    /// Provides extension methods to inject capability into a [Site]
    /// </summary>
    public static class SiteExtensions
    {
        /// <summary>
        /// Adds or Updates an existing Custom Action [ScriptSrc] into the [Site] Custom Actions
        /// </summary>
        /// <param name="site"></param>
        /// <param name="customactionname"></param>
        /// <param name="customactionurl"></param>
        /// <param name="sequence"></param>
        public static void AddOrUpdateCustomActionLink(this Site site, string customactionname, string customactionurl, int sequence)
        {
            var sitecustomActions = site.GetCustomActions();
            UserCustomAction cssAction = null;
            if (site.CustomActionExists(customactionname))
            {
                cssAction = sitecustomActions.FirstOrDefault(fod => fod.Name == customactionname);
            }
            else
            {
                // Build a custom action to write a link to our new CSS file
                cssAction = site.UserCustomActions.Add();
                cssAction.Name = customactionname;
                cssAction.Location = "ScriptLink";
            }

            cssAction.Sequence = sequence;
            cssAction.ScriptSrc = customactionurl;
            cssAction.Update();
            site.Context.ExecuteQueryRetry();
        }

        /// <summary>
        /// Adds or Updates an existing Custom Action [ScriptBlock] into the [Site] Custom Actions
        /// </summary>
        /// <param name="site"></param>
        /// <param name="customactionname"></param>
        /// <param name="customActionBlock"></param>
        /// <param name="sequence"></param>
        public static void AddOrUpdateCustomActionLinkBlock(this Site site, string customactionname, string customActionBlock, int sequence)
        {
            var sitecustomActions = site.GetCustomActions();
            UserCustomAction cssAction = null;
            if (site.CustomActionExists(customactionname))
            {
                cssAction = sitecustomActions.FirstOrDefault(fod => fod.Name == customactionname);
            }
            else
            {
                // Build a custom action to write a link to our new CSS file
                cssAction = site.UserCustomActions.Add();
                cssAction.Name = customactionname;
                cssAction.Location = "ScriptLink";
            }

            cssAction.Sequence = sequence;
            cssAction.ScriptBlock = customActionBlock;
            cssAction.Update();
            site.Context.ExecuteQueryRetry();
        }

        /// <summary>
        /// Will remove the custom action if one exists
        /// </summary>
        /// <param name="site"></param>
        /// <param name="customactionname"></param>
        public static void RemoveCustomActionLink(this Site site, string customactionname)
        {
            if (site.CustomActionExists(customactionname))
            {
                var cssAction = site.GetCustomActions().FirstOrDefault(fod => fod.Name == customactionname);
                site.DeleteCustomAction(cssAction.Id);
            }
        }
    }
}
