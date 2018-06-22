using InfrastructureAsCode.Core.Utilities;
using Microsoft.SharePoint.Client;
using System.Linq;

namespace InfrastructureAsCode.Core.Extensions
{
    public static class ClientExtensions
    {
        /// <summary>
        /// Load the collection and execute retry
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="collection"></param>
        /// <returns></returns>
        public static T Load<T>(this T collection) where T : ClientObjectCollection
        {
            if (collection.ServerObjectIsNull == null || collection.ServerObjectIsNull == true)
            {
                collection.Context.Load(collection);
                collection.Context.ExecuteQueryRetry();
                return collection;
            }
            else
            {
                return collection;
            }
        }

        /// <summary>
        /// Take the URL and clean it
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        public static string EnsureTrailingSlashLowered(this string url)
        {
            var surl = url;
            if (!string.IsNullOrEmpty(surl))
            {
#if !NETSTANDARD2_0
                surl = TokenHelper.EnsureTrailingSlash(url.Trim().ToLower());
#endif
            }
            return surl;
        }

        /// <summary>
        /// Returns the [Tenant] admin URI from the established context
        /// </summary>
        /// <param name="context"></param>
        /// <returns></returns>
        public static string GetTenantAdminUri(this ClientContext context)
        {
            var uri = new System.Uri(context.Url);
            return uri.GetTenantAdminUri();
        }
    }
}
