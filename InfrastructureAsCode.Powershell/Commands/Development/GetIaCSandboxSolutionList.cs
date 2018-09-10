using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Powershell.Commands.Development
{
    using Microsoft.SharePoint.Client;
    using InfrastructureAsCode.Core.Extensions;
    using InfrastructureAsCode.Core.Models;
    using InfrastructureAsCode.Powershell.Commands.Base;
    using Microsoft.Online.SharePoint.TenantAdministration;
    using OfficeDevPnP.Core.Entities;
    using OfficeDevPnP.Core.Utilities;

    /// <summary>
    /// Opens tenant and enumerates all sites pulling the sandbox solution set
    /// </summary>
    [Cmdlet(VerbsCommon.Get, "EPASandboxSolutionList")]
    public class GetIaCSandboxSolutionList : IaCAdminCmdlet
    {

        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();


            // Log file for output with the current time
            var date = DateTime.Now;

            LogVerbose("---------------------------------------------------------------------------");
            LogVerbose("|               Get List of Sandbox solutions from tenant                 |");
            LogVerbose("---------------------------------------------------------------------------");

            LogVerbose("Admin Site URL: {0}", this.ClientContext.Url);

            var urls = new List<SPOSiteCollectionModel>();
            var sandboxedSolutions = new List<object>();

            //Retrieve all site collection infos
            ClientContext.Load(TenantContext);
            ClientContext.ExecuteQueryRetry();

            var rootSiteUrl = base.TenantContext.RootSiteUrl;

            LogVerbose("Querying Site Collections; this might take a few minutes");
            urls = GetSiteCollections();
      

            // Counter for the progress tracking
            var count = 1;

            // Scan through the sites and output content from solution gallery (exists in root site of site collection)
            foreach (var site in urls)
            {
                try
                {
                    SetSiteAdmin(site.Url, CurrentUserName, true);

                    // Connect to root web
                    var context = this.ClientContext.Clone(site.Url);
                    var web = context.Web;
                    context.Load(web);
                    context.ExecuteQueryRetry();

                    // Get the sandboxed solution gallery - Catalog code 121
                    var solutionGallery = web.GetCatalog(121);
                    context.Load(solutionGallery);
                    context.ExecuteQueryRetry();

                    // Get items from the solution gallery
                    var query = CamlQuery.CreateAllItemsQuery();
                    var items = solutionGallery.GetItems(query);
                    context.Load(items);
                    context.ExecuteQueryRetry();


                    if (items.Count > 0)
                    {

                        // List the sandbox solutions
                        foreach (var item in items)
                        {
                            // Resolve status of the solution - 1=Activate, 0=not active
                            var statusField = item.RetrieveListItemValue("Status").ToBoolean(false);
                            var status = 0;
                            if (statusField)
                            {
                                status = 1;
                            }

                            // Output to console
                            var metaInfo = item.RetrieveListItemValue("MetaInfo");
                            var author = item.RetrieveListItemUserValue("Author");
                            var solutionFile = item.RetrieveListItemValue("FileLeafRef");
                            var created = item.RetrieveListItemValue("Created").ToDateTime();
                            var solutionSize = item.RetrieveListItemValueAsLookup("SMTotalSize");
                            var solutionFileCount = item.RetrieveListItemValueAsLookup("SMTotalFileCount");
                            var solutionFileType = item.RetrieveListItemValue("File_x0020_Type");

                            LogVerbose("{0},{1},{2},{3}", site.Url, solutionFile, author.LookupValue, created, status);

                            // Output report in format, which can be imported to excel
                            sandboxedSolutions.Add(new
                            {
                                url = site.Url,
                                fileLeaf = solutionFile,
                                author = author.LookupValue,
                                created = created,
                                status = status,
                                fileType = solutionFileType,
                                metadata = metaInfo
                            });
                        }

                    }
                    //# Output to file right next to script location
                    LogVerbose("Scanning site collections Status {0} %{1} complete", site.Url, Convert.ToInt32(count * 100.0 / urls.Count));
                    // Tracking progress
                    count++;

                    SetSiteAdmin(site.Url, CurrentUserName, false);
                }
                catch (Exception ex)
                {
                    // Possible public site exception handler
                    LogWarning("Exception occurred!");
                    LogError(ex, "Failed in URL Parsing");
                }
            }

            LogVerbose("");
            LogVerbose("----------------------------------------------");
            LogVerbose("|               List produced                |");
            LogVerbose("----------------------------------------------");
            LogVerbose("Next steps: Import output file generated folder with ps1 to excel for analyses");
            LogVerbose("Columns are Site URL, Sandbox solution name, Author and Created");

            WriteObject(sandboxedSolutions, true);
        }
    }


}
