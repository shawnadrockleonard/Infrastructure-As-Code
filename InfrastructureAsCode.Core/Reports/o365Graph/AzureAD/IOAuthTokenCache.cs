using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Reports.o365Graph.AzureAD
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
        Task<string> AccessToken();

        /// <summary>
        /// If the token is no longer fresh it will claim a new token
        /// </summary>
        /// <returns>Authentication Result which contains a Token and ExpiresOn</returns>
        Task<AuthenticationResult> AccessTokenResult();

        /// <summary>
        /// Acquires AuthenticationResult without asking for user credential.
        /// </summary>
        /// <returns></returns>
        Task<AuthenticationResult> GetTokenForAadGraph();

        /// <summary>
        ///     Acquires security token from the authority using an authorization code previously
        ///     received. This method does not lookup token cache, but stores the result in it,
        /// </summary>
        /// <param name="code"></param>
        /// <param name="redirect_uri"></param>
        /// <returns></returns>
        Task RedeemAuthCodeForAadGraph(string code, string redirect_uri);

        /// <summary>
        /// Clears the user token cache
        /// </summary>
        void Clear();
    }
}