using InfrastructureAsCode.Core;
using InfrastructureAsCode.Core.Models;
using InfrastructureAsCode.Core.oAuth;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Powershell.CmdLets
{
    /// <summary>
    /// Holds all of the information about the current Azure AD Connection and OAuth 2.0 Access Token
    /// </summary>
    public class AzureADALv1Connection
    {
        public static AzureADALv1Connection CurrentConnection { get; set; }

        /// <summary>
        /// Holds the OAuth 2.0 Authentication Result
        /// </summary>
        public AuthenticationResult AuthenticationResult
        {
            get
            {
                return GetTokenAsyncResult();
            }
        }


        public SPOAddInKeys AddInCredentials { get; protected set; }


        public AzureADConfig AzureADCredentials { get; protected set; }


        public AzureADTokenCache AzureADCache { get; protected set; }


        private readonly ITraceLogger _iLogger;

        /// <summary>
        /// Initialize the Azure AD Connection with config and diagnostics
        /// </summary>
        /// <param name="azureADCredentials"></param>
        /// <param name="traceLogger"></param>
        public AzureADALv1Connection(AzureADConfig azureADCredentials, ITraceLogger traceLogger)
        {
            _iLogger = traceLogger;
            AzureADCache = new AzureADTokenCache(azureADCredentials, traceLogger);
            AzureADCredentials = azureADCredentials;
        }

        /// <summary>
        /// Initiates a blocker and waites for a Async thread to complete
        /// </summary>
        /// <returns></returns>
        internal AuthenticationResult GetTokenAsyncResult()
        {
            AuthenticationResult authenticationResult = null;
            try
            {
                var asyncFunction = AzureADCache.TryGetAccessTokenResultAsync(string.Empty);

                asyncFunction.Wait();

                authenticationResult = asyncFunction.Result;
            }
            catch (Exception ex)
            {
                _iLogger.LogError(ex, $"Claiming Azure AD Token Failed {ex.Message}");
            }

            if (authenticationResult == null || string.IsNullOrEmpty(authenticationResult.AccessToken))
            {
                throw new ArgumentNullException("authenticationResult");
            }

            return authenticationResult;
        }
    }
}
