using Microsoft.SharePoint.Client;
using System;
using System.Management.Automation;
using System.Security;
using InfrastructureAsCode.Core.Enums;
using InfrastructureAsCode.Core.oAuth;
using InfrastructureAsCode.Core.Models;

namespace InfrastructureAsCode.Powershell.CmdLets
{
    /// <summary>
    /// Provides a connection object for operating on SharePoint instance
    /// </summary>
    public class SPIaCConnection
    {
        private ClientContext _initialContext;

        public static SPIaCConnection CurrentConnection { get; set; }

        public ConnectionType ConnectionType { get; protected set; }

        public int MinimalHealthScore { get; protected set; }

        public int RetryCount { get; protected set; }

        public int RetryWait { get; protected set; }

        public PSCredential PSCredential { get; protected set; }

        public string Url { get; protected set; }

        public ClientContext Context { get; set; }

        /// <summary>
        /// Contains the Azure AD Config to claim authorization
        /// </summary>
        public IAzureADConfig AzureConfig { get; set; }

        public SPOAddInKeys AddInCredentials { get; set; }


        /// <summary>
        /// Initializes the OnlineConnection for connecting via Federation/Integrated/OAuth
        /// </summary>
        /// <param name="context"></param>
        /// <param name="connectionType"></param>
        /// <param name="minimalHealthScore"></param>
        /// <param name="retryCount"></param>
        /// <param name="retryWait"></param>
        /// <param name="credential"></param>
        /// <param name="url"></param>
        public SPIaCConnection(ClientContext context, ConnectionType connectionType, int minimalHealthScore, int retryCount, int retryWait, PSCredential credential, string url)
        {
            if (context == null)
                throw new ArgumentNullException("context");
            Context = context;
            _initialContext = context;
            ConnectionType = connectionType;
            MinimalHealthScore = minimalHealthScore;
            RetryCount = retryCount;
            RetryWait = retryWait;
            PSCredential = credential;
            Url = url;
        }

        public void RestoreCachedContext()
        {
            Context = _initialContext;
        }

        internal void CacheContext()
        {
            _initialContext = Context;
        }

        /// <summary>
        /// Returns the user who initiated the connection
        /// </summary>
        /// <returns></returns>
        public string GetActiveUsername()
        {
            if (CurrentConnection != null)
            {
                return CurrentConnection.PSCredential.UserName;
            }
            return string.Empty;
        }

        /// <summary>
        /// Returns the active credentials for the SPO connection
        /// </summary>
        /// <returns></returns>
        public PSCredential GetActiveCredentials()
        {
            if (CurrentConnection != null)
            {
                return CurrentConnection.PSCredential;
            }
            return null;
        }
    }
}
