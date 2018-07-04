using InfrastructureAsCode.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Powershell.CmdLets
{
    /// <summary>
    /// Azure AD v1 EndPoint 
    ///     >>> Base class for all the Microsoft Graph related cmdlets
    /// </summary>
    public abstract class AzureADALv1Cmdlet : ExtendedPSCmdlet
    {
        /// <summary>
        /// The current ADAL v1 connection with Access Tokens
        /// </summary>
        public AzureADALv1Connection Connection
        {
            get { return AzureADALv1Connection.CurrentConnection; }
        }


        public String AccessToken
        {
            get
            {
                if (Connection != null)
                {
                    return (Connection.AuthenticationResult.AccessToken);
                }
                else
                {
                    WriteError(new ErrorRecord(new InvalidOperationException("NoAzureADAccessToken"), "NO_OAUTH_TOKEN", ErrorCategory.ConnectionError, null));
                    return (null);
                }
            }
        }

        internal ITraceLogger Logger { get; private set; }


        protected override void OnBeginInitialize()
        {
            base.OnBeginInitialize();

            Logger = new DefaultUsageLogger(LogVerbose, LogWarning, LogError);


            if (Connection == null || Connection.AuthenticationResult == null)
            {
                throw new InvalidOperationException("NoAzureADAccessToken");
            }
        }


        protected override void ProcessRecord()
        {
            ExecuteCmdlet();
        }
    }
}