using InfrastructureAsCode.Core.oAuth;
using InfrastructureAsCode.Core.Utilities;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Utilities;
using System;
using System.Management.Automation;
using System.Reflection;
using System.Threading;
using System.Xml.Linq;
using Resources = InfrastructureAsCode.Core.Properties.Resources;

namespace InfrastructureAsCode.Powershell.Commands.Base
{
    public abstract class IaCCmdlet : ExtendedPSCmdlet, IIaCCmdlet
    {
        public IaCCmdlet() : base()
        {

        }

        public ClientContext ClientContext
        {
            get { return SPIaCConnection.CurrentConnection.Context; }
        }

        public IAzureADConfig AzureADConfig
        {
            get { return SPIaCConnection.CurrentConnection.AzureConfig; }
        }


        /// <summary>
        /// The base URI for the SP Site or Tenant
        /// </summary>
        internal string BaseUri { get; private set; }

        /// <summary>
        /// Represents the claim identifier prefix
        /// </summary>
        internal const string ClaimIdentifier = "i:0#.f|membership";



        /// <summary>
        /// Initializers the logger from the cmdlet
        /// </summary>
        protected override void OnBeginInitialize()
        {
            base.OnBeginInitialize();

            if (SPIaCConnection.CurrentConnection == null)
            {
                throw new InvalidOperationException(Resources.NoConnection);
            }

            if (ClientContext == null)
            {
                throw new InvalidOperationException(Resources.NoConnection);
            }

            Uri uri = new Uri(this.ClientContext.Url);
            var urlParts = uri.Authority.Split(new[] { '.' });
            BaseUri = string.Format("https://{0}.{1}.{2}", urlParts[0], urlParts[1], urlParts[2]);

        }

        /// <summary>
        /// Process SPO HealthCheck and validation context
        /// </summary>
        protected override void ProcessRecord()
        {
            try
            {
                if (SPIaCConnection.CurrentConnection.MinimalHealthScore != -1)
                {
                    int healthScore = Utility.GetHealthScore(SPIaCConnection.CurrentConnection.Url);
                    if (healthScore <= SPIaCConnection.CurrentConnection.MinimalHealthScore)
                    {
                        ExecuteCmdlet();
                    }
                    else
                    {
                        if (SPIaCConnection.CurrentConnection.RetryCount != -1)
                        {
                            int retry = 1;
                            while (retry <= SPIaCConnection.CurrentConnection.RetryCount)
                            {
                                WriteWarning(string.Format(Resources.Retry0ServerNotHealthyWaiting1seconds, retry, SPIaCConnection.CurrentConnection.RetryWait, healthScore));
                                Thread.Sleep(SPIaCConnection.CurrentConnection.RetryWait * 1000);
                                healthScore = Utility.GetHealthScore(SPIaCConnection.CurrentConnection.Url);
                                if (healthScore <= SPIaCConnection.CurrentConnection.MinimalHealthScore)
                                {
                                    ExecuteCmdlet();
                                    break;
                                }
                                retry++;
                            }
                            if (retry > SPIaCConnection.CurrentConnection.RetryCount)
                            {
                                WriteError(new ErrorRecord(new Exception(Resources.HealthScoreNotSufficient), "HALT", ErrorCategory.OpenError, null));
                            }
                        }
                        else
                        {
                            WriteError(new ErrorRecord(new Exception(Resources.HealthScoreNotSufficient), "HALT", ErrorCategory.OpenError, null));
                        }
                    }
                }
                else
                {
                    ExecuteCmdlet();
                }
            }
            catch (Exception ex)
            {
                SPIaCConnection.CurrentConnection.RestoreCachedContext();
                System.Diagnostics.Trace.TraceError("Cmdlet Exception {0}", ex.Message);
                if (!this.Stopping)
                {
                    LogError(ex, "Stack Trace {0}", ex.StackTrace);
                }
            }
        }

        /// <summary>
        /// internal member to hold the current user
        /// </summary>
        private string _currentUserInProcess = string.Empty;

        /// <summary>
        /// this should be valid based on pre authentication checks
        /// </summary>
        protected virtual string CurrentUserName
        {
            get
            {
                if (string.IsNullOrEmpty(_currentUserInProcess))
                {
                    try
                    {
                        _currentUserInProcess = SPIaCConnection.CurrentConnection.GetActiveUsername();
                    }
                    catch (Exception ex)
                    {
                        LogError(ex, "Failed to retrieve the current context credential for the Cached Entity");
                    }
                }
                return _currentUserInProcess;
            }
        }

        /// <summary>
        /// internal member to hold the current network credentials
        /// </summary>
        private System.Net.NetworkCredential _currentNetworkInProcess = null;

        /// <summary>
        /// this should the current network credentials
        /// </summary>
        protected virtual System.Net.NetworkCredential CurrentNetworkCredential
        {
            get
            {
                if (_currentNetworkInProcess == null)
                {
                    try
                    {
                        var tmpcurrentUserInProcess = SPIaCConnection.CurrentConnection.GetActiveCredentials();
                        _currentNetworkInProcess = tmpcurrentUserInProcess.GetNetworkCredential();
                    }
                    catch (Exception ex)
                    {
                        LogError(ex, "Failed to retrieve the current context credential for the Cached Entity");
                    }
                }

                return _currentNetworkInProcess;
            }
        }
    }
}
