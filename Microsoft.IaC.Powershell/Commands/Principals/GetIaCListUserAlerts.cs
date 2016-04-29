using IaC.Core.Models;
using IaC.Powershell;
using IaC.Powershell.CmdLets;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;

namespace IaC.Powershell.Commands.Principals
{
    [Cmdlet(VerbsCommon.Get, "IaCListUserAlerts")]
    [CmdletHelp("Opens a web request and queries the users alerts", Category = "Principals")]
    public class GetIaCListUserAlerts : IaCCmdlet
    {
        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 1)]
        public string ListTitle;

        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();

            var models = new List<SPUserAlertDefinition>();
            try
            {
                var creds = SPIaCConnection.CurrentConnection.GetActiveCredentials();
                var newcreds = new System.Net.NetworkCredential(creds.UserName, creds.Password);
                var spourl = new Uri(this.ClientContext.Url);
                var spocreds = new Microsoft.SharePoint.Client.SharePointOnlineCredentials(creds.UserName, creds.Password);
                var spocookies = spocreds.GetAuthenticationCookie(spourl);

                var spocontainer = new System.Net.CookieContainer();
                spocontainer.SetCookies(spourl, spocookies);

                var ws = new IaC.Core.com.sharepoint.useralerts.Alerts();
                ws.Url = string.Format("{0}/_vti_bin/Alerts.asmx", spourl.AbsoluteUri);
                ws.Credentials = newcreds;
                ws.CookieContainer = spocontainer;
                var alerts = ws.GetAlerts();

                LogVerbose("User {0} webId:{1} has {2} alerts configured", alerts.CurrentUser, alerts.AlertWebId, alerts.Alerts.Count());
                foreach (var alertItem in alerts.Alerts)
                {
                    var model = new SPUserAlertDefinition()
                    {
                        CurrentUser = alerts.CurrentUser,
                        WebId = alerts.AlertWebId,
                        WebTitle = alerts.AlertWebTitle,
                        AlertForTitle = alertItem.AlertForTitle,
                        AlertForUrl = alertItem.AlertForUrl,
                        EventType = alertItem.EventType,
                        Id = alertItem.Id
                    };
                    models.Add(model);
                    LogVerbose("Alert {0} Active:{1} EventType:{2} Id:{3}", alertItem.AlertForUrl, alertItem.Active, alertItem.EventType, alertItem.Id);
                }


                models.ForEach(alert => WriteObject(alert));
            }
            catch (Exception ex)
            {
                LogError(ex, "Failed in GetListItemCount for Library {0}", ListTitle);
            }
        }
    }
}
