using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
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
        private readonly ITraceLogger _iLogger;

        /// <summary>
        /// Represents the token to be used during authentication
        /// </summary>
        internal static AuthenticationResult AuthenticationToken { get; private set; }

        public AzureADTokenCache(string authenticationEndpoint, IAzureADConfig aadConfig, ITraceLogger iLogger)
        {
            _aadConfig = aadConfig;
            _authContext = new AuthenticationContext(authenticationEndpoint);
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
        /// <param name="redirectUri">(OPTIONAL) a redirect to the resource URI</param>
        /// <returns></returns>
        async public Task<string> AccessTokenAsync(string redirectUri)
        {
            var result = await AccessTokenResultAsync(redirectUri);
            return result.AccessToken;
        }

        /// <summary>
        /// Will request the token, if the cache has expired, will throw an exception and request a new auth cache token and attempt to return it
        /// </summary>
        /// <param name="redirectUri">(OPTIONAL) a redirect to the resource URI</param>
        /// <returns>Return an Authentication Result which contains the Token/Refresh Token</returns>
        async public Task<AuthenticationResult> TryGetAccessTokenResultAsync(string redirectUri)
        {
            AuthenticationResult token = null;
            var cleanToken = false;

            try
            {
                token = await AccessTokenResultAsync(redirectUri);
                cleanToken = true;
            }
            catch (Exception ex)
            {
                _iLogger.LogError(ex, "AdalCacheException: {0}", ex.Message);
            }

            if (!cleanToken)
            {
                await RedeemAuthCodeForAadGraphAsync(string.Empty, redirectUri);
                token = await AccessTokenResultAsync(redirectUri);
            }

            return token;
        }

        /// <summary>
        /// Validate the current token in the cache
        /// </summary>
        /// <param name="redirectUri">(OPTIONAL) a redirect to the resource URI</param>
        /// <returns></returns>
        async public Task<AuthenticationResult> AccessTokenResultAsync(string redirectUri)
        {
            if (AuthenticationToken == null
                || AuthenticationToken.ExpiresOn <= DateTimeOffset.Now)
            {
                AuthenticationToken = await GetTokenForAadGraphAsync(redirectUri);
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

        /// <summary>
        /// Redeem the Authentication Code and return the AuthenticationResult
        /// </summary>
        /// <param name="redirectUri">(OPTIONAL) a redirect to the resource URI</param>
        /// <returns></returns>
        async public Task<AuthenticationResult> GetTokenForAadGraphAsync(string redirectUri)
        {
            await RedeemAuthCodeForAadGraphAsync(string.Empty, redirectUri);
            return AuthenticationToken;
        }

        /// <summary>
        /// Returns Azure AD Code response after initial signin and acceptance of the profile access
        /// </summary>
        /// <param name="code">The Azure AD Code response after initial signin and acceptance of the profile access</param>
        /// <param name="redirectUri">(OPTIONAL) a redirect to the resource URI</param>
        /// <returns></returns>
        async public Task RedeemAuthCodeForAadGraphAsync(string code, string redirectUri)
        {
            // Failed to retrieve, reup the token
            redirectUri = (string.IsNullOrEmpty(redirectUri) ? GetRedirectUri() : redirectUri);

            // Redeem the auth code and cache the result for later use.
            if (!_aadConfig.IsCertificateAuth())
            {
                var _appCredentials = new ClientCredential(_aadConfig.ClientId, _aadConfig.ClientSecret);
                var result = await _authContext.AcquireTokenAsync(redirectUri, _appCredentials);
                AuthenticationToken = result;
            }
            else
            {
                var _certificateCredentials = new ClientAssertionCertificate(_aadConfig.ClientId, _aadConfig.Certificate);
                var result = await _authContext.AcquireTokenAsync(redirectUri, _certificateCredentials);
                AuthenticationToken = result;
            }
        }
    }
}
