using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.oAuth
{
    /// <summary>
    /// OAuth interface for claiming Tokens
    /// </summary>
    public interface IOAuthTokenCache
    {
        /// <summary>
        /// Return the Redirect URI from the AzureAD Config
        /// </summary>
        string GetRedirectUri();

        /// <summary>
        /// If the token is no longer fresh it will claim a new token
        /// </summary>
        /// <returns>Access Token as a string</returns>
        Task<string> AccessTokenAsync(string redirectUri);

        /// <summary>
        /// Retreive the access token from the ClientCredentials
        /// </summary>
        /// <param name="redirectUri">(TOPIONAL) a redirect to the resource URI</param>
        /// <returns></returns>
        /// <remarks>Will handle automatic refresh of the tokens</remarks>
        Task<AuthenticationResult> TryGetAccessTokenResultAsync(string redirectUri);

        /// <summary>
        /// If the token is no longer fresh it will claim a new token
        /// </summary>
        /// <returns>Authentication Result which contains a Token and ExpiresOn</returns>
        Task<AuthenticationResult> AccessTokenResultAsync(string redirectUri);

        /// <summary>
        /// Acquires AuthenticationResult without asking for user credential.
        /// </summary>
        /// <returns></returns>
        Task<AuthenticationResult> GetTokenForAadGraphAsync(string redirectUri);

        /// <summary>
        /// Clears the user token cache
        /// </summary>
        void Clear();
    }
}