using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.oAuth
{
    /// <summary>
    /// Represents a generic token cache to pull Tokens or Refresh Tokens
    /// </summary>
    public class AzureADTokenCache : IOAuthTokenCache
    {
        private readonly IAzureADConfig _aadConfig;
        private readonly AuthenticationContext _authContext;
        private readonly ClientCredential _appCredentials;
        private readonly ITraceLogger _iLogger;

        /// <summary>
        /// Represents the token to be used during authentication
        /// </summary>
        internal static AuthenticationResult AuthenticationToken { get; private set; }

        public AzureADTokenCache(IAzureADConfig aadConfig, ITraceLogger iLogger)
        {
            _aadConfig = aadConfig;
            _authContext = new AuthenticationContext(string.Format(AzureADConstants.AuthorityTenantFormat, _aadConfig.TenantDomain));
            _appCredentials = new ClientCredential(_aadConfig.ClientId, _aadConfig.ClientSecret);
            _iLogger = iLogger;
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
        async public Task<string> AccessTokenAsync()
        {
            var result = await AccessTokenResultAsync();
            return result.AccessToken;
        }

        /// <summary>
        /// Will request the token, if the cache has expired, will throw an exception and request a new auth cache token and attempt to return it
        /// </summary>
        /// <param name="redirectUri">(OPTIONAL) a redirect to the resource URI</param>
        /// <returns>Return an Authentication Result which contains the Token/Refresh Token</returns>
        async public Task<AuthenticationResult> TryGetAccessTokenResultAsync(string redirectUri)
        {
            AuthenticationResult token = null; var cleanToken = false;

            try
            {
                token = await AccessTokenResultAsync();
                cleanToken = true;
            }
            catch (Exception ex)
            {
                _iLogger.LogError(ex, "AdalCacheException: {0}", ex.Message);
            }

            if (!cleanToken)
            {
                // Failed to retrieve, reup the token
                redirectUri = (string.IsNullOrEmpty(redirectUri) ? GetRedirectUri() : redirectUri);
                await RedeemAuthCodeForAadGraphAsync(string.Empty, redirectUri);
                token = await AccessTokenResultAsync();
            }

            return token;
        }

        /// <summary>
        /// Validate the current token in the cache
        /// </summary>
        /// <returns></returns>
        async public Task<AuthenticationResult> AccessTokenResultAsync()
        {
            if (AuthenticationToken == null
                || AuthenticationToken.ExpiresOn <= DateTimeOffset.Now)
            {
                AuthenticationToken = await GetTokenForAadGraphAsync();
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

        async public Task<AuthenticationResult> GetTokenForAadGraphAsync()
        {
            await RedeemAuthCodeForAadGraphAsync(string.Empty, _aadConfig.RedirectUri);
            return AuthenticationToken;
        }

        async public Task RedeemAuthCodeForAadGraphAsync(string code, string resource_uri)
        {
            // Redeem the auth code and cache the result in the db for later use.
            var result = await _authContext.AcquireTokenAsync(resource_uri, _appCredentials);
            AuthenticationToken = result;
        }
    }
}
