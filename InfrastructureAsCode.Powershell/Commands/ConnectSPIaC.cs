using InfrastructureAsCode.Core.Extensions;
using InfrastructureAsCode.Powershell.CmdLets;
using InfrastructureAsCode.Powershell.PipeBinds;
using OfficeDevPnP.Core.Utilities;
using System;
using System.Management.Automation;
using System.Reflection;
using System.Security;
using Resources = InfrastructureAsCode.Core.Properties.Resources;

namespace InfrastructureAsCode.Powershell.Commands
{
    /*
        Examples:

        This will prompt for username and password and creates a context for the other PowerShell commands to use.
        Connect-SPIaC -Url https://yourtenant.sharepoint.com -Credentials (Get-Credential)
    
        This will use the current user credentials and connects to the server specified by the Url parameter.
        Connect-SPIaC -Url http://yourlocalserver -CurrentCredentials
        
        This will use credentials from the Windows Credential Manager, as defined by the label 'O365Creds'
        Connect-SPIaC -Url http://yourlocalserver -Credentials 'O365Creds'
    */
    [Cmdlet(VerbsExtended.Connect, "SPIaC", SupportsShouldProcess = false)]
    [CmdletHelp("Connects to a SharePoint site and creates an in-memory context", DetailedDescription = "If no credentials have been specified, and the CurrentCredentials parameter has not been specified, you will be prompted for credentials.", Category = "Base Cmdlets")]
    public class ConnectSPIaC : ExtendedPSCmdlet
    {
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterAttribute.AllParameterSets, ValueFromPipeline = true, HelpMessage = "The Url of the site collection to connect to.")]
        public string Url;

        [Parameter(Mandatory = false, ParameterSetName = "Main", HelpMessage = "Credentials of the user to connect with. Either specify a PSCredential object or a string. In case of a string value a lookup will be done to the Windows Credential Manager for the correct credentials.")]
        public CredentialPipeBind Credentials;

        [Parameter(Mandatory = false, ParameterSetName = "Main", HelpMessage = "If you want to connect with the current user credentials")]
        public SwitchParameter CurrentCredentials;

        [Parameter(Mandatory = false, ParameterSetName = ParameterAttribute.AllParameterSets, HelpMessage = "Specifies a minimal server healthscore before any requests are executed.")]
        public int MinimalHealthScore = -1;

        [Parameter(Mandatory = false, ParameterSetName = ParameterAttribute.AllParameterSets, HelpMessage = "Defines how often a retry should be executed if the server healthscore is not sufficient.")]
        public int RetryCount = -1;

        [Parameter(Mandatory = false, ParameterSetName = ParameterAttribute.AllParameterSets, HelpMessage = "Defines how many seconds to wait before each retry. Default is 5 seconds.")]
        public int RetryWait = 5;

        [Parameter(Mandatory = false, ParameterSetName = ParameterAttribute.AllParameterSets, HelpMessage = "The request timeout. Default is 180000")]
        public int RequestTimeout = 1800000;

        [Parameter(Mandatory = false, ParameterSetName = "Token")]
        public string Realm { get; set; }

        [Parameter(Mandatory = true, ParameterSetName = "Token")]
        public string AppId { get; set; }

        [Parameter(Mandatory = true, ParameterSetName = "Token")]
        public string AppSecret { get; set; }

        [Parameter(Mandatory = true, HelpMessage = "The AAD where the O365 app is registered. Eg.: contoso.com, or contoso.onmicrosoft.com.", ParameterSetName = "Token")]
        public string AppDomain { get; set; }

        [Parameter(Mandatory = true, HelpMessage = "The URI of the resource to query", ParameterSetName = "Token")]
        public string ResourceUri { get; set; }


        [Parameter(Mandatory = true, ParameterSetName = "UserCache")]
        public string UserName { get; set; }

        /// <summary>
        /// Represents a parameter to pull from the stored credentials
        /// </summary>
        [Parameter(Mandatory = true, ParameterSetName = "CredentialCache")]
        public string CredentialName { get; set; }

        /// <summary>
        /// Remove the need to check if this is a tenant client context
        /// </summary>
        [Parameter(Mandatory = false, ParameterSetName = ParameterAttribute.AllParameterSets)]
        public SwitchParameter SkipTenantAdminCheck;



        public override void ExecuteCmdlet()
        {
            PSCredential creds = null;
            if (Credentials != null)
            {
                creds = Credentials.Credential;
            }


            if (ParameterSetName == "Token")
            {
                SPIaCConnection.CurrentConnection = SPIaCConnectionHelper.InstantiateSPOnlineConnection(new Uri(Url), Realm, AppId, AppSecret, AppDomain, ResourceUri, Host, MinimalHealthScore, RetryCount, RetryWait, RequestTimeout, SkipTenantAdminCheck);
            }
            else if (ParameterSetName == "CredentialCache")
            {
                var genericcreds = CredentialManager.GetCredential(CredentialName);
                creds = new PSCredential(genericcreds.UserName, genericcreds.SecurePassword);
                SPIaCConnection.CurrentConnection = SPIaCConnectionHelper.InstantiateSPOnlineConnection(new Uri(Url), creds, Host, CurrentCredentials, MinimalHealthScore, RetryCount, RetryWait, RequestTimeout, SkipTenantAdminCheck);
            }
            else if (ParameterSetName == "UserCache")
            {
                var boolSaveToDisk = false;
                var runningDirectory = this.SessionState.Path.CurrentFileSystemLocation;
                var userPasswordConfig = string.Format("{0}\\{1}.pswd", runningDirectory, UserName).Replace("\\", @"\");
                if (System.IO.File.Exists(userPasswordConfig))
                {
                    var encryptedUserPassword = System.IO.File.ReadAllText(userPasswordConfig);
                    var encryptedSecureString = encryptedUserPassword.ConvertToSecureString();

                    if (!CurrentCredentials && creds == null)
                    {
                        if (encryptedSecureString == null || encryptedSecureString.Length <= 0)
                        {
                            boolSaveToDisk = true;
                            creds = Host.UI.PromptForCredential(Resources.EnterYourCredentials, "", UserName, "");
                        }
                        else
                        {
                            creds = new PSCredential(this.UserName, encryptedSecureString);
                        }
                    }
                }
                else
                {
                    // the password was not encrypted
                    if (!CurrentCredentials && creds == null)
                    {
                        boolSaveToDisk = true;
                        creds = Host.UI.PromptForCredential(Resources.EnterYourCredentials, "", UserName, "");
                    }
                }

                var initializedConnection = SPIaCConnectionHelper.InstantiateSPOnlineConnection(new Uri(Url), creds, Host, CurrentCredentials, MinimalHealthScore, RetryCount, RetryWait, RequestTimeout, SkipTenantAdminCheck);
                if (initializedConnection == null)
                {
                    throw new Exception(string.Format("Error establishing a connection to {0}.  Check the diagnostic logs.", Url));
                }

                SPIaCConnection.CurrentConnection = initializedConnection;

                if (boolSaveToDisk)
                {
                    var encryptedSecureString = creds.Password.ConvertFromSecureString();
                    System.IO.File.WriteAllText(userPasswordConfig, encryptedSecureString);
                }
            }
            else
            {
                if (!CurrentCredentials && creds == null)
                {
                    creds = Host.UI.PromptForCredential(Resources.EnterYourCredentials, "", "", "");
                }

                LogVerbose("Received credentials for {0} user", creds.UserName);

                var initializedConnection = SPIaCConnectionHelper.InstantiateSPOnlineConnection(new Uri(Url), creds, Host, CurrentCredentials, MinimalHealthScore, RetryCount, RetryWait, RequestTimeout, SkipTenantAdminCheck);
                if (initializedConnection == null)
                {
                    throw new Exception(string.Format("Error establishing a connection to {0}.  Check the diagnostic logs.", Url));
                }

                SPIaCConnection.CurrentConnection = initializedConnection;

                LogVerbose("Processed credentials for {0} user", creds.UserName);

            }
        }

    }
}
