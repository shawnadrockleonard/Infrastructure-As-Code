using InfrastructureAsCode.Core.Reports.o365Graph.AzureAD;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Reports.o365Graph
{
    public class ReportingStream
    {
        #region Internals

        /// <summary>
        /// Represents the Graph API endpoints
        /// </summary>
        /// <remarks>Of note this is a BETA inpoint as these APIs are in Preview</remarks>
        public static string DefaultServiceEndpointUrl = "https://graph.microsoft.com/beta/reports/{0}({1})/content";

        public IAzureADConfig ADConfig { get; private set; }

        public ITraceLogger Logger { get; private set; }

        public IOAuthTokenCache OAuthCache { get; private set; }

        public QueryFilter ServiceQuery { get; private set; }

        #endregion Privates

        #region Properties

        public ITraceLogger TraceLogger
        {
            get
            {
                return this.Logger;
            }
        }

        public string GraphUrl { get; set; }

        #endregion

        public ReportingStream(QueryFilter serviceQuery, IAzureADConfig config, ITraceLogger logger)
            : this(DefaultServiceEndpointUrl, serviceQuery, config, logger)
        {
        }

        public ReportingStream(string url, QueryFilter serviceQuery, IAzureADConfig config, ITraceLogger logger)
        {
            this.GraphUrl = url;
            this.ADConfig = config;
            this.Logger = logger;
            this.ServiceQuery = serviceQuery;
            this.OAuthCache = new AzureADTokenCache(config);
        }

        /// <summary>
        /// Will request the token, if the cache has expired, will throw an exception and request a new auth cache token and attempt to return it
        /// </summary>
        /// <returns>Return an Authentication Result which contains the Token/Refresh Token</returns>
        private async Task<AuthenticationResult> GetAccessTokenResult()
        {
            AuthenticationResult token = null; var cleanToken = false;

            try
            {
                token = await OAuthCache.AccessTokenResult();
                cleanToken = true;
            }
            catch (Exception ex)
            {
                Logger.LogError(ex, "AdalCacheException: {0}", ex.Message);
            }

            if (!cleanToken)
            {
                // Failed to retrieve, reup the token
                var redirectUri = OAuthCache.GetRedirectUri();
                await OAuthCache.RedeemAuthCodeForAadGraph(string.Empty, redirectUri);
                token = await OAuthCache.AccessTokenResult();
            }

            return token;
        }

        /// <summary>
        /// Initiates a blocker and waites for a Async thread to complete
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="asyncFunction"></param>
        /// <returns></returns>
        private T GetAsyncResult<T>(Task<T> asyncFunction)
        {
            asyncFunction.Wait();
            return asyncFunction.Result;
        }

        private Uri BuildServiceQuery()
        {
            var serviceFullUrl = ServiceQuery.ToUrl(GraphUrl);
            Logger.LogInformation("Request URL : {0}", serviceFullUrl);
            return serviceFullUrl;
        }

        public void RetrieveData()
        {
            ReportVisitor visitor = new DefaultReportVisitor(Logger);
            RetrieveData(visitor);
        }

        /// <summary>
        /// Builds the URI from the Reporting types and returns the streamer
        /// </summary>
        /// <param name="visitor"></param>
        /// <returns></returns>
        public void RetrieveData(ReportVisitor visitor)
        {

            var Token = GetAsyncResult(GetAccessTokenResult());

            var serviceFullUrl = BuildServiceQuery();
            var webRequest = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(serviceFullUrl);
            webRequest.Method = "GET";
            webRequest.ContentType = "application/json";
            webRequest.Headers.Add(System.Net.HttpRequestHeader.Authorization, Token.CreateAuthorizationHeader());

            var webResponse = webRequest.GetResponse();
            using (Stream webStream = webResponse.GetResponseStream())
            {
                using (StreamReader responseReader = new StreamReader(webStream))
                {
                    try
                    {
                        if (responseReader != null)
                        {
                            visitor.ProcessReport(responseReader);
                        }
                        else
                        {
                            throw new Exception("Response content is Null");
                        }
                    }
                    catch (HttpRequestException hex)
                    {
                        Logger.LogError(hex, "HTTP Failed to query URI {0}", serviceFullUrl);
                        throw hex;
                    }
                    catch (Exception ex)
                    {
                        Logger.LogError(ex, "Generic Failed to query URI {0}", serviceFullUrl);
                        throw ex;
                    }
                }
            }
        }
    }
}
