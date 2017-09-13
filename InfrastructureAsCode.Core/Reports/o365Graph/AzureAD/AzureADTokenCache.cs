using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Reports.o365Graph.AzureAD
{
    public class AzureADTokenCache : IOAuthTokenCache
    {
        private readonly IAzureADConfig _aadConfig;
        private readonly AuthenticationContext _authContext;
        private readonly ClientCredential _appCredentials;

        /// <summary>
        /// Represents the token to be used during authentication
        /// </summary>
        internal static AuthenticationResult AuthenticationToken { get; private set; }

        public AzureADTokenCache(IAzureADConfig aadConfig)
        {
            _aadConfig = aadConfig;
            _authContext = new AuthenticationContext(string.Format(AzureADConstants.AuthorityTenantFormat, _aadConfig.TenantDomain));
            _appCredentials = new ClientCredential(_aadConfig.ClientId, _aadConfig.ClientSecret);
        }

        /// <summary>
        /// Return the Redirect URI from the AzureAD Config
        /// </summary>
        public string GetRedirectUri()
        {
            return _aadConfig.RedirectUri.ToString();
        }

        /// <summary>
        /// Validate the current token in the cache
        /// </summary>
        /// <returns></returns>
        async public Task<string> AccessToken()
        {
            var result = await AccessTokenResult();
            return result.AccessToken;
        }

        /// <summary>
        /// Validate the current token in the cache
        /// </summary>
        /// <returns></returns>
        async public Task<AuthenticationResult> AccessTokenResult()
        {
            if (AuthenticationToken == null
                || AuthenticationToken.ExpiresOn <= DateTimeOffset.Now)
            {
                AuthenticationToken = await GetTokenForAadGraph();
            }

            return AuthenticationToken;
        }

        /// <summary>
        /// clean up the db
        /// </summary>
        public void Clear()
        {
            AuthenticationToken = null;
        }

        async public Task<AuthenticationResult> GetTokenForAadGraph()
        {
            await RedeemAuthCodeForAadGraph(string.Empty, _aadConfig.RedirectUri);
            return AuthenticationToken;
        }

        async public Task RedeemAuthCodeForAadGraph(string code, string resource_uri)
        {
            // Redeem the auth code and cache the result in the db for later use.
            var result = await _authContext.AcquireTokenAsync(resource_uri, _appCredentials);
            AuthenticationToken = result;
        }
    }
}
