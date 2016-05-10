using InfrastructureAsCode.Powershell;
using InfrastructureAsCode.Powershell.CmdLets;
using InfrastructureAsCode.Core.Extensions;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using InfrastructureAsCode.Core.Models;

namespace InfrastructureAsCode.Powershell.Commands.Development
{
    /// <summary>
    /// This command will query the App Catalog site parsing the App Manifests for permissions and status
    /// </summary>
    /// <remarks>
    /// The app manifest typical includes the following
    ///    <AppPrincipal xmlns = "http://schemas.microsoft.com/sharepoint/2012/app/manifest" >
    ///     <RemoteWebApplication ClientId="17b6a002-cfbe-403d-ba89-b6c9f7a18773" />
    ///     </AppPrincipal>
    ///     <AppPermissionRequests xmlns = "http://schemas.microsoft.com/sharepoint/2012/app/manifest" >
    ///     <AppPermissionRequest Scope="http://sharepoint/content/sitecollection/web" Right="Manage" />
    ///     <AppPermissionRequest Scope = "http://sharepoint/social/tenant" Right="Read" />
    ///     <AppPermissionRequest Scope = "http://sharepoint/content/tenant" Right="Read" />
    ///    </AppPermissionRequests>
    /// </remarks>
    [Cmdlet(VerbsCommon.Get, "IaCAppPermissions")]
    [CmdletHelp("Opens the app catalog and scans apps for permissions", Category = "Development")]
    public class GetIaCAppPermissions : SPOAdminCmdlet
    {
        /// <summary>
        /// A collection of internal names to retreive and dump to a txt file
        /// </summary>
        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 1)]
        public string AppCatalogUrl { get; set; }

        /// <summary>
        /// Execute the command and query the app catalog scanning for App permissions
        /// </summary>
        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();

            var models = new List<SPAppManifestModel>();

            try
            {
                // Set connection to app catalog
                var appCatalogListName = "App Requests";
                var appCatalogSiteContext = this.ClientContext.Clone(AppCatalogUrl);
                var appCatalogWeb = appCatalogSiteContext.Web;
                var appCatalogList = appCatalogWeb.Lists.GetByTitle(appCatalogListName);
                appCatalogSiteContext.Load(appCatalogWeb);
                appCatalogSiteContext.Load(appCatalogList);
                appCatalogSiteContext.ExecuteQueryRetry();


                LogVerbose("App Catalog list: {0}", appCatalogList.Title);

                ListItemCollectionPosition itemPosition = null;

                var tmpFalg = true;
                var viewFields = new string[]
                {
                "Title",
                "AppPublisher",
                "AppRequester",
                "AppRequestJustification",
                "AppRequestIsSiteLicense",
                "AppRequestPermissionXML",
                "AppRequestStatus",
                "AssetID"
                };
                var camlQuery = CamlQuery.CreateAllItemsQuery(50, viewFields);

                while (tmpFalg)
                {
                    camlQuery.ListItemCollectionPosition = itemPosition;
                    ListItemCollection spListItems = appCatalogList.GetItems(camlQuery);

                    appCatalogSiteContext.Load(spListItems);
                    appCatalogSiteContext.ExecuteQueryRetry();
                    itemPosition = spListItems.ListItemCollectionPosition;
                    var tmpTitle = string.Empty;

                    foreach (var item in spListItems)
                    {
                        var itemTitle = item.RetrieveListItemValue("Title");
                        if (!itemTitle.Equals(tmpTitle, StringComparison.CurrentCultureIgnoreCase))
                        {
                            var model = new SPAppManifestModel()
                            {
                                Title = itemTitle,
                                AssetId = item.RetrieveListItemValue("AssetID"),
                                AppRequestStatus = item.RetrieveListItemValue("AppRequestStatus"),
                                AppRequestIsSiteLicense = item.RetrieveListItemValue("AppRequestIsSiteLicense")
                            };

                            LogVerbose("Scanning {0}", itemTitle);

                            var xmlFromField = item.RetrieveListItemValue("AppRequestPermissionXML");
                            if (!string.IsNullOrEmpty(xmlFromField))
                            {
                                var appXml = XDocument.Parse(string.Format("<appxml xmlns=\"{1}\">{0}</appxml>", xmlFromField, "http://schemas.microsoft.com/sharepoint/2012/app/manifest"), LoadOptions.None);
                                var appType = appXml.Root.GetType();
                                var appRequestsName = XName.Get("AppPermissionRequests", "http://schemas.microsoft.com/sharepoint/2012/app/manifest");
                                var appRequests = appXml.Root.Element(appRequestsName);
                                if (appRequests != null)
                                {
                                    var appRequestName = XName.Get("AppPermissionRequest", "http://schemas.microsoft.com/sharepoint/2012/app/manifest");
                                    var appRequestItems = appRequests.Elements(appRequestName);
                                    foreach (var appItem in appRequestItems)
                                    {
                                        var appRight = appItem.Attribute("Right").Value;
                                        var appScope = appItem.Attribute("Scope").Value;
                                        model.AppPermissions.Add(new SPAppScopePermissionModel()
                                        {
                                            AppRights = appRight,
                                            AppScope = appScope
                                        });
                                    }
                                }
                            }

                            models.Add(model);
                        }
                    }

                    if (itemPosition == null)
                    {
                        break;
                    }

                }

                // Write the app information to the console
                models.ForEach(app => WriteObject(app));
            }
            catch (Exception ex)
            {
                LogError(ex, "Failed in QueryAppMonitor for {0}", this.AppCatalogUrl);
            }
        }
        
    }
}
