using InfrastructureAsCode.Core.Constants;
using InfrastructureAsCode.Core.oAuth;
using Microsoft.Identity.Client;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.HttpServices
{
    public class GraphHttpHelper
    {
        public const String GraphBaseUrl = "https://graph.microsoft.com";
        public const String MicrosoftGraphV1BaseUri = GraphBaseUrl + "/v1.0/";
        public const String MicrosoftGraphBetaBaseUri = GraphBaseUrl + "/beta/";

        internal IAppConfig ConfigKeys { get; private set; }

        /// <summary>
        /// Initialize from an App.Config
        /// </summary>
        public GraphHttpHelper()
        {
            ConfigKeys = (new AppConfigManager()).GetApplicationConfig();
        }

        /// <summary>
        /// Feed the App Config domain model
        /// </summary>
        /// <param name="appConfig"></param>
        public GraphHttpHelper(IAppConfig appConfig)
        {
            ConfigKeys = appConfig;
        }

        /// <summary>
        /// Will acquire a token from the confidential client via the "Delegated" access or Elevated permission set
        /// </summary>
        /// <returns></returns>
        public async Task<string> GetGraphDaemonAccessTokenAsync()
        {
            var tenantId = ConfigKeys.TenantId;
            var iss = string.Format(ConstantsAuthentication.TenantAuthority, tenantId);
            var lowerscopes = new string[] { ConstantsAuthentication.MSGraphScope };

            try
            {
                ConfidentialClientApplication daemonClient = new ConfidentialClientApplication(ConfigKeys.MSALClientID, iss, ConfigKeys.PostLogoutRedirectURI, new ClientCredential(ConfigKeys.MSALClientSecret), null, new TokenCache());
                AuthenticationResult authResult = await daemonClient.AcquireTokenForClientAsync(lowerscopes);
                return authResult.AccessToken;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.TraceError($"Failed to claim credentials {iss} w/ {ex.Message}");
                return null;
            }
        }


        /// <summary>
        /// This helper method makes an HTTP request and eventually returns a result
        /// </summary>
        /// <param name="httpMethod">The HTTP method for the request</param>
        /// <param name="requestUrl">The URL of the request</param>
        /// <param name="accept">The content type of the accepted response</param>
        /// <param name="content">The content of the request</param>
        /// <param name="contentType">The content  type of the request</param>
        /// <param name="resultPredicate">The predicate to retrieve the result, if any</param>
        /// <param name="accessToken">The AccessToken to use for the request</param>
        /// <typeparam name="TResult">The type of the result, if any</typeparam>
        /// <returns>The value of the result, if any</returns>
        private static TResult MakeHttpRequest<TResult>(
            String httpMethod,
            String requestUrl,
            String accept = null,
            Object content = null,
            String contentType = null,
            String accessToken = null,
            Func<HttpResponseMessage, TResult> resultPredicate = null)
        {
            // Prepare the variable to hold the result, if any
            TResult result = default;

            // Get the OAuth Access Token
            if (String.IsNullOrEmpty(accessToken))
            {
                throw new ArgumentException("Invalid value for accessToken", nameof(accessToken));
            }
            else
            {
                // If we have the token, then handle the HTTP request
                using HttpClientHandler handler = new HttpClientHandler
                {
                    AllowAutoRedirect = true
                };
                HttpClient httpClient = new HttpClient(handler, true);

                // Set the Authorization Bearer token
                httpClient.DefaultRequestHeaders.Authorization =
                    new AuthenticationHeaderValue("Bearer", accessToken);

                // If there is an accept argument, set the corresponding HTTP header
                if (!String.IsNullOrEmpty(accept))
                {
                    httpClient.DefaultRequestHeaders.Accept.Clear();
                    httpClient.DefaultRequestHeaders.Accept.Add(
                        new MediaTypeWithQualityHeaderValue(accept));
                }

                // Prepare the content of the request, if any
                HttpContent requestContent = null;
                if (content is System.IO.Stream streamContent)
                {
                    requestContent = new StreamContent(streamContent);
                    requestContent.Headers.ContentType = new MediaTypeHeaderValue(contentType);
                }
                else
                {
                    requestContent =
                        (content != null) ?
                        new StringContent(JsonConvert.SerializeObject(content,
                            Formatting.None,
                            new JsonSerializerSettings
                            {
                                NullValueHandling = NullValueHandling.Ignore,
                                ContractResolver = new CamelCasePropertyNamesContractResolver(),
                            }),
                        Encoding.UTF8, contentType) :
                        null;
                }

                // Prepare the HTTP request message with the proper HTTP method
                HttpRequestMessage request = new HttpRequestMessage(
                    new HttpMethod(httpMethod), requestUrl);

                // Set the request content, if any
                if (requestContent != null)
                {
                    request.Content = requestContent;
                }

                // Fire the HTTP request
                HttpResponseMessage response = httpClient.SendAsync(request).Result;

                if (response.IsSuccessStatusCode)
                {
                    // If the response is Success and there is a
                    // predicate to retrieve the result, invoke it
                    if (resultPredicate != null)
                    {
                        result = resultPredicate(response);
                    }
                }
                else
                {
                    throw new ApplicationException(
                        String.Format("Exception while invoking endpoint {0}.", requestUrl),
#if !NETSTANDARD2_0
                            new System.Web.HttpException(
                            (Int32)response.StatusCode,
                            response.Content.ReadAsStringAsync().Result));
#else
                            new Exception(response.Content.ReadAsStringAsync().Result));
#endif
                }
            }

            return (result);
        }

        /// <summary>
        /// This helper method makes an HTTP GET request and returns the result as a String
        /// </summary>
        /// <param name="requestUrl">The URL of the request</param>
        /// <param name="accessToken">The AccessToken to use for the request</param>
        /// <returns>The String value of the result</returns>
        public static String MakeGetRequestForString(String requestUrl, String accessToken = null)
        {
            return (MakeHttpRequest<String>("GET",
                requestUrl,
                accessToken: accessToken,
                resultPredicate: r => r.Content.ReadAsStringAsync().Result));
        }

        /// <summary>
        /// This helper method makes an HTTP GET request and returns the result as a String
        /// </summary>
        /// <param name="requestUrl">The URL of the request</param>
        /// <param name="accept">The accept header for the response</param>
        /// <param name="accessToken">The AccessToken to use for the request</param>
        /// <returns>The Stream  of the result</returns>
        public static System.IO.Stream MakeGetRequestForStream(String requestUrl, String accept, String accessToken = null)
        {
            return (MakeHttpRequest<System.IO.Stream>("GET",
                requestUrl,
                accessToken: accessToken,
                resultPredicate: r => r.Content.ReadAsStreamAsync().Result));
        }

        /// <summary>
        /// This helper method makes an HTTP POST request without a response
        /// </summary>
        /// <param name="requestUrl">The URL of the request</param>
        /// <param name="content">The content of the request</param>
        /// <param name="contentType">The content/type of the request</param>
        /// <param name="accessToken">The AccessToken to use for the request</param>
        public static void MakePostRequest(String requestUrl,
            Object content = null,
            String contentType = null,
            String accessToken = null)
        {
            MakeHttpRequest<String>("POST",
                requestUrl,
                content: content,
                contentType: contentType,
                accessToken: accessToken);
        }

        /// <summary>
        /// This helper method makes an HTTP POST request and returns the result as a String
        /// </summary>
        /// <param name="requestUrl">The URL of the request</param>
        /// <param name="content">The content of the request</param>
        /// <param name="contentType">The content/type of the request</param>
        /// <param name="accessToken">The AccessToken to use for the request</param>
        /// <returns>The String value of the result</returns>
        public static String MakePostRequestForString(String requestUrl,
            Object content = null,
            String contentType = null,
            String accessToken = null)
        {
            return (MakeHttpRequest<String>("POST",
                requestUrl,
                content: content,
                contentType: contentType,
                accessToken: accessToken,
                resultPredicate: r => r.Content.ReadAsStringAsync().Result));
        }

        /// <summary>
        /// This helper method makes an HTTP PUT request without a response
        /// </summary>
        /// <param name="requestUrl">The URL of the request</param>
        /// <param name="content">The content of the request</param>
        /// <param name="contentType">The content/type of the request</param>
        /// <param name="accessToken">The AccessToken to use for the request</param>
        public static void MakePutRequest(String requestUrl,
            Object content = null,
            String contentType = null,
            String accessToken = null)
        {
            MakeHttpRequest<String>("PUT",
                requestUrl,
                content: content,
                contentType: contentType,
                accessToken: accessToken);
        }

        /// <summary>
        /// This helper method makes an HTTP PATCH request and returns the result as a String
        /// </summary>
        /// <param name="requestUrl">The URL of the request</param>
        /// <param name="content">The content of the request</param>
        /// <param name="contentType">The content/type of the request</param>
        /// <param name="accessToken">The AccessToken to use for the request</param>
        /// <returns>The String value of the result</returns>
        public static String MakePatchRequestForString(String requestUrl,
            Object content = null,
            String contentType = null,
            String accessToken = null)
        {
            return (MakeHttpRequest<String>("PATCH",
                requestUrl,
                content: content,
                contentType: contentType,
                accessToken: accessToken,
                resultPredicate: r => r.Content.ReadAsStringAsync().Result));
        }

        /// <summary>
        /// This helper method makes an HTTP DELETE request
        /// </summary>
        /// <param name="requestUrl">The URL of the request</param>
        /// <param name="accessToken">The AccessToken to use for the request</param>
        /// <returns>The String value of the result</returns>
        public static void MakeDeleteRequest(String requestUrl,
            String accessToken = null)
        {
            MakeHttpRequest<String>("DELETE", requestUrl, accessToken: accessToken);
        }

    }
}
