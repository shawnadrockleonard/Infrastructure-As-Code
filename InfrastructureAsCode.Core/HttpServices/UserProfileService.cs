using InfrastructureAsCode.Core.Extensions;
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

        public OfficeDevPnP.Core.UPAWebService.UserProfileService OWService { get; private set; }

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

            var trailingSiteUrl = siteUrl.EnsureTrailingSlashLowered();

            OWService = new OfficeDevPnP.Core.UPAWebService.UserProfileService
            {
                Url = $"{trailingSiteUrl}_vti_bin/userprofileservice.asmx",
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
                OWService.CookieContainer = cookieContainer;
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
                OWService.Dispose();
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
