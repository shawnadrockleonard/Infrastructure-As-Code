using InfrastructureAsCode.Core.Reports.o365Graph.AzureAD;
using InfrastructureAsCode.Core.Reports.o365Graph.TenantReport;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Reports.o365Graph
{
    public class ReportingStream
    {
        #region Internals

        /// <summary>
        /// Collection of Azure AD settings required to claim tokens
        /// </summary>
        public IAzureADConfig ADConfig { get; private set; }

        /// <summary>
        /// Diagnostic Logger for event listeners
        /// </summary>
        public ITraceLogger Logger { get; private set; }

        /// <summary>
        /// OAuth cache class for retreiving tokens
        /// </summary>
        public IOAuthTokenCache OAuthCache { get; private set; }

        #endregion


        /// <summary>
        /// Initialize the Graph API Executor with Azure AD Config settings and the Diagnostic Logger
        /// </summary>
        /// <param name="config"></param>
        /// <param name="logger"></param>
        public ReportingStream(IAzureADConfig config, ITraceLogger logger)
        {
            this.ADConfig = config;
            this.Logger = logger;
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

        /// <summary>
        /// Builds the URI from the Reporting types and returns the streamer
        /// </summary>
        /// <param name="serviceFullUrl">The full URL to the Graph API</param>
        /// <param name="maxAttempts">total number of attempts before proceeding</param>
        /// <param name="backoffIntervalInSeconds">wait interval (in seconds) before retry</param>
        /// <returns></returns>
        internal string ExecuteResponse(Uri serviceFullUrl, int maxAttempts = 3, int backoffIntervalInSeconds = 6)
        {
            var resultResponse = string.Empty;
            var retry = false;
            var retryAttempts = 0;
            // Reset the Default backoff in Seconds
            var graphBackoffInterval = backoffIntervalInSeconds;

            do
            {
                try
                {
                    retry = false;
                    graphBackoffInterval = backoffIntervalInSeconds;

                    // Retreive the Access Token
                    var Token = GetAsyncResult(GetAccessTokenResult());

                    // Establish the HTTP Request
                    var webRequest = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(serviceFullUrl);
                    webRequest.Method = "GET";
                    webRequest.ContentType = "application/json";
                    webRequest.Headers.Add(System.Net.HttpRequestHeader.Authorization, Token.CreateAuthorizationHeader());
                    this.Logger.LogInformation("Executing {0}", serviceFullUrl);
                    using (var webResponse = (System.Net.HttpWebResponse)webRequest.GetResponse())
                    {
                        using (var webStream = webResponse.GetResponseStream())
                        {
                            if (webStream != null)
                            {
                                using (var responseReader = new StreamReader(webStream))
                                {
                                    resultResponse = responseReader.ReadToEnd();
                                }
                            }
                            else
                            {
                                throw new Exception("Response content is Null");
                            }
                        }
                    }
                }
                catch (System.Net.WebException wex)
                {
                    // Check if request was throttled - http status code 429
                    // Check is request failed due to server unavailable - http status code 503
                    if (wex.Response is HttpWebResponse response &&
                        (response.StatusCode == (HttpStatusCode)429 || response.StatusCode == (HttpStatusCode)503))
                    {
                        // Extract the Retry-After throttling suggestion
                        var graphApiRetrySeconds = response.GetResponseHeader("Retry-After");
                        if (!string.IsNullOrEmpty(graphApiRetrySeconds)
                            && int.TryParse(graphApiRetrySeconds, out int headergraphBackoffInterval))
                        {
                            if(headergraphBackoffInterval <= 0)
                            {
                                graphBackoffInterval = backoffIntervalInSeconds;
                            }
                            else
                            {
                                graphBackoffInterval = headergraphBackoffInterval;
                            }
                        }
                        var backoffSpan = new TimeSpan(0, 0, 0, graphBackoffInterval, 0);

                        Logger.LogWarning("Microsoft Graph API => exceeded usage limits. Iteration => {1} Sleeping for {0} seconds before retrying..", backoffSpan.Seconds, retryAttempts);
                        
                        //Add delay for retry
                        Task.Delay(backoffSpan).Wait();

                        //Add to retry count and check max attempts.
                        retryAttempts++;
                        retry = (retryAttempts < maxAttempts);
                    }
                    else
                    {
                        Logger.LogError(wex, "HTTP Failed to query URI {0} exception: {1}", serviceFullUrl, wex.ToString());
                        throw;
                    }
                }
                catch (Exception ex)
                {
                    Logger.LogWarning("Generic Failed to query URI {0} => {1}", serviceFullUrl, ex.Message);
                    throw;
                }
            }
            while (retry);


            return resultResponse;
        }

        /// <summary>
        /// Builds the URI from the Reporting types and returns the string
        /// </summary>
        /// <param name="serviceQuery">The GraphAPI URI builder object with specific settings</param>
        /// <param name="maxAttempts">total number of attempts before proceeding</param>
        /// <param name="backoffIntervalInSeconds">wait interval (in seconds) before retry</param>
        /// <returns></returns>
        public string RetrieveData(QueryFilter serviceQuery, int maxAttempts = 3, int backoffIntervalInSeconds = 6)
        {
            var serviceFullUrl = serviceQuery.ToUrl();

            var result = ExecuteResponse(serviceFullUrl, maxAttempts, backoffIntervalInSeconds);
            return result;
        }

        /// <summary>
        /// Builds the URI from the Reporting types and returns the Deserialized Objects
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="serviceQuery">The GraphAPI URI builder object with specific settings</param>
        /// <param name="maxAttempts">total number of attempts before proceeding</param>
        /// <param name="backoffIntervalInSeconds">wait interval (in seconds) before retry</param>
        /// <returns>A deserialized collection of objects</returns>
        public JSONAuditCollection<T> RetrieveData<T>(QueryFilter serviceQuery, int maxAttempts = 3, int backoffIntervalInSeconds = 6) where T : class
        {
            JSONAuditCollection<T> objects = new JSONAuditCollection<T>();
            var serviceFullUrl = serviceQuery.ToUrl();
            var lastUri = serviceFullUrl;

            while (true)
            {
                var result = ExecuteResponse(lastUri, maxAttempts, backoffIntervalInSeconds);
                if (string.IsNullOrEmpty(result))
                {
                    break;
                }

                var items = JsonConvert.DeserializeObject<JSONAuditCollection<T>>(result);
                objects.value.AddRange(items.value);
                if (string.IsNullOrEmpty(items.NextLink))
                {
                    // last in the set
                    break;
                }
                lastUri = new Uri(items.NextLink);
            }
            return objects;
        }

        /// <summary>
        /// Builds the URI from the Reporting types and returns the streamer
        /// </summary>
        /// <param name="serviceQuery">The GraphAPI URI builder object with specific settings</param>
        /// <param name="maxAttempts">total number of attempts before proceeding</param>
        /// <param name="backoffIntervalInSeconds">wait interval (in seconds) before retry</param>
        /// <returns>An open Text reader which should be disposed</returns>
        public TextReader RetrieveDataAsStream(QueryFilter serviceQuery, int maxAttempts = 3, int backoffIntervalInSeconds = 6)
        {
            var serviceFullUrl = serviceQuery.ToUrl();

            var result = ExecuteResponse(serviceFullUrl, maxAttempts, backoffIntervalInSeconds);
            TextReader textReader = new StringReader(result);
            return textReader;
        }


    }
}
