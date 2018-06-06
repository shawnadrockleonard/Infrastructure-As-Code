using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.HttpServices
{
    public class UserProfileService : IDisposable
    {
        private bool _disposed;

        public OfficeDevPnP.Core.UPAWebService.UserProfileService ows { get; private set; }

        public UserProfileService()
        {

        }

        public UserProfileService(ClientContext context, string siteUrl = "") : this()
        {
            if (string.IsNullOrEmpty(siteUrl))
            {
                if (!context.Site.IsPropertyAvailable(sctx => sctx.Url))
                {
                    context.Site.EnsureProperties(sctx => sctx.Url);
                    siteUrl = context.Site.Url;
                }
            }

            ows = new OfficeDevPnP.Core.UPAWebService.UserProfileService
            {
                Url = string.Format("{0}/_vti_bin/userprofileservice.asmx", siteUrl),
                Credentials = context.Credentials,
                UseDefaultCredentials = false
            };


            if (context.Credentials is SharePointOnlineCredentials)
            {
                var spourl = new Uri(siteUrl);
                var newcreds = (SharePointOnlineCredentials)context.Credentials;
                var spocookies = newcreds.GetAuthenticationCookie(spourl);

                var cookieContainer = new System.Net.CookieContainer();
                cookieContainer.SetCookies(spourl, spocookies);
                ows.CookieContainer = cookieContainer;
            }
        }


        private void ThrowIfDisposed()
        {
            if (_disposed)
            {
                throw new ObjectDisposedException(GetType().Name);
            }
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing && !_disposed)
            {
                ows.Dispose();
            }
            _disposed = true;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
    }
}
