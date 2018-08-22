using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Powershell.Commands.ListItems
{
    using Microsoft.SharePoint.Client;
    using InfrastructureAsCode.Powershell.PipeBinds;
    using InfrastructureAsCode.Powershell.Commands.Base;
    using InfrastructureAsCode.Core.Models;
    using InfrastructureAsCode.Powershell;


    [Cmdlet(VerbsCommon.Get, "IaCListItemsREST")]
    [CmdletHelp("Opens a web request and queries the specified list via REST", Category = "ListItems")]
    public class GetIaCListItemsREST : IaCCmdlet
    {
        /// <summary>
        /// The display name for the list or library to query
        /// </summary>
        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 1)]
        public ListPipeBind ListTitle { get; set; }

        /// <summary>
        /// A collection of internal names to retreive and dump to a txt file
        /// </summary>
        [Parameter(Mandatory = false, ValueFromPipeline = true, Position = 2)]
        public int? Throttle { get; set; }

        /// <summary>
        /// Initialize a default value if not specified by the user
        /// </summary>
        protected override void OnBeginInitialize()
        {
            base.OnBeginInitialize();

            if (!Throttle.HasValue)
            {
                Throttle = 200;
            }
        }

        /// <summary>
        /// Execute the REST API querying the list with paging
        /// </summary>
        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();

            Collection<SPListItemDefinition> results = new Collection<SPListItemDefinition>();

            try
            {
                var creds = SPIaCConnection.CurrentConnection.GetActiveCredentials();
                var spourl = new Uri(this.ClientContext.Url);
                var spocreds = new Microsoft.SharePoint.Client.SharePointOnlineCredentials(creds.UserName, creds.Password);
                var spocookies = spocreds.GetAuthenticationCookie(spourl);
                var spocontainer = new System.Net.CookieContainer();
                spocontainer.SetCookies(spourl, spocookies);

                // region Consume the web service
                var ListService = string.Format("{0}/_api/web/lists/getByTitle('{1}')/ItemCount", this.ClientContext.Url, this.ListTitle);
                var webRequest = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(ListService);
                webRequest.Credentials = new System.Net.NetworkCredential(creds.UserName, creds.Password);
                webRequest.Method = "GET";
                webRequest.Accept = "application/json;odata=verbose";
                webRequest.CookieContainer = spocontainer;

                var webResponse = webRequest.GetResponse();
                using (Stream webStream = webResponse.GetResponseStream())
                {
                    using (StreamReader responseReader = new StreamReader(webStream))
                    {
                        string response = responseReader.ReadToEnd();
                        var jobj = JObject.Parse(response);
                        var itemCount = jobj["d"]["ItemCount"];
                        LogVerbose("ItemCount:{0}", itemCount);
                    }
                }

                var successFlag = true;
                ListService = string.Format("{0}/_api/web/lists/getByTitle('{1}')/items?$top={2}", this.ClientContext.Url, this.ListTitle, this.Throttle);
                while (successFlag)
                {
                    LogVerbose("Paging:{0}", ListService);
                    successFlag = false;
                    webRequest = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(ListService);
                    webRequest.Credentials = new System.Net.NetworkCredential(creds.UserName, creds.Password);
                    webRequest.Method = "GET";
                    webRequest.Accept = "application/json;odata=verbose";
                    webRequest.CookieContainer = spocontainer;

                    webResponse = webRequest.GetResponse();
                    using (Stream webStream = webResponse.GetResponseStream())
                    {
                        using (StreamReader responseReader = new StreamReader(webStream))
                        {
                            string response = responseReader.ReadToEnd();
                            var jobj = JObject.Parse(response);
                            var jarr = (JArray)jobj["d"]["results"];
                            var jnextpage = jobj["d"]["__next"];

                            foreach (JObject j in jarr)
                            {
                                LogVerbose("ItemID:{0}", j["Id"]);
                                var newitem = new SPListItemDefinition()
                                {
                                    Title = j["Title"].ToObject<string>(),
                                    Id = j["Id"].ToObject<int>()
                                };
                                results.Add(newitem);
                            }

                            if (jnextpage != null && !String.IsNullOrEmpty(jnextpage.ToString()))
                            {
                                successFlag = true;
                                ListService = jnextpage.ToString();
                            }
                        }
                    }
                }

                WriteObject(results, true);
            }
            catch (Exception ex)
            {
                LogError(ex, "Failed in GetListItemCount for Library {0}", ListTitle);
            }
        }

        /// <summary>
        /// Retreives the internal column name value for the list item
        /// </summary>
        /// <param name="j"></param>
        /// <param name="columnName"></param>
        /// <returns></returns>
        private string GetColumnValue(JObject j, string columnName)
        {
            JToken rtypeval = null; var rval = string.Empty;
            if (j.TryGetValue(columnName, out rtypeval))
            {
                rval = rtypeval.ToString();
            }
            return rval;
        }
    }
}
