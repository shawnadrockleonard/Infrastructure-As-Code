using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Extensions
{
    /// <summary>
    /// Provides extensions to handle URLs
    /// </summary>
    public static class UriExtensions
    {
        /// <summary>
        /// Pulls the [Tenant] URL from a Context URI
        /// </summary>
        /// <param name="siteUri"></param>
        /// <returns></returns>
        public static string GetTenantAdminUri(this Uri siteUri)
        {
            string scenarioUrl = $"{siteUri.Scheme}://{siteUri.DnsSafeHost}";
            if (scenarioUrl.IndexOf(@"-admin", StringComparison.CurrentCultureIgnoreCase) == -1)
            {
                scenarioUrl = scenarioUrl.Replace(".sharepoint.com", "-admin.sharepoint.com");
            }

            return scenarioUrl;
        }

        /// <summary>
        /// Retreives the Realm or Tenant ID from the client.svc
        /// </summary>
        /// <param name="targetApplicationUri"></param>
        /// <returns></returns>
        public static string GetRealmFromTargetUrl(this Uri targetApplicationUri)
        {
            WebRequest request = WebRequest.Create(targetApplicationUri + "/_vti_bin/client.svc");
            request.Headers.Add("Authorization: Bearer ");

            try
            {
                using (request.GetResponse())
                {
                }
            }
            catch (WebException e)
            {
                if (e.Response == null)
                {
                    return null;
                }

                string bearerResponseHeader = e.Response.Headers["WWW-Authenticate"];
                if (string.IsNullOrEmpty(bearerResponseHeader))
                {
                    return null;
                }

                const string bearer = "Bearer realm=\"";
                int bearerIndex = bearerResponseHeader.IndexOf(bearer, StringComparison.Ordinal);
                if (bearerIndex < 0)
                {
                    return null;
                }

                int realmIndex = bearerIndex + bearer.Length;

                if (bearerResponseHeader.Length >= realmIndex + 36)
                {
                    string targetRealm = bearerResponseHeader.Substring(realmIndex, 36);

                    Guid realmGuid;

                    if (Guid.TryParse(targetRealm, out realmGuid))
                    {
                        return targetRealm;
                    }
                }
            }
            return null;
        }
    }
}
