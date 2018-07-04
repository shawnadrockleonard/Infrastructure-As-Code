using InfrastructureAsCode.Core.Enums;
using InfrastructureAsCode.Core.Extensions;
using InfrastructureAsCode.Core.Models;
using InfrastructureAsCode.Core.oAuth;
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using System;
using System.Management.Automation;
using System.Management.Automation.Host;
using System.Net;
using Resources = InfrastructureAsCode.Core.Properties.Resources;

namespace InfrastructureAsCode.Powershell.CmdLets
{
    /// <summary>
    /// Helper class to instantiate the proper authentication manager for onpremise, online
    /// </summary>
    internal class SPIaCConnectionHelper
    {
        static SPIaCConnectionHelper()
        {
        }

        internal static SPIaCConnection InstantiateSPOnlineConnection(Uri url, string realm, string clientId, string clientSecret, string appDomain, string resourceUri, PSHost host, int minimalHealthScore, int retryCount, int retryWait, int requestTimeout, bool skipAdminCheck = false)
        {
            var authManager = new OfficeDevPnP.Core.AuthenticationManager();
            if (realm == null)
            {
                realm = url.GetRealmFromTargetUrl();
            }

            var context = authManager.GetAppOnlyAuthenticatedContext(url.ToString(), realm, clientId, clientSecret);
            context.ApplicationName = Resources.ApplicationName;
            context.RequestTimeout = requestTimeout;

            var connectionType = ConnectionType.OnPrem;
            if (url.Host.ToUpperInvariant().EndsWith("SHAREPOINT.COM"))
            {
                connectionType = ConnectionType.O365;
            }

            if (skipAdminCheck == false)
            {
                if (IsTenantAdminSite(context))
                {
                    connectionType = ConnectionType.TenantAdmin;
                }
            }

            var connection = new SPIaCConnection(context, connectionType, minimalHealthScore, retryCount, retryWait, null, url.ToString())
            {
                AddInCredentials = new SPOAddInKeys()
                {
                    AppId = clientId,
                    AppKey = clientSecret,
                    Realm = realm
                }
            };
            return connection;
        }

        internal static SPIaCConnection InstantiateSPOnlineConnection(Uri url, PSCredential credentials, PSHost host, bool currentCredentials, int minimalHealthScore, int retryCount, int retryWait, int requestTimeout, bool skipAdminCheck = false)
        {
            ClientContext context = new ClientContext(url.AbsoluteUri)
            {
                ApplicationName = Resources.ApplicationName,
                RequestTimeout = requestTimeout
            };

            if (!currentCredentials)
            {
                try
                {
                    SharePointOnlineCredentials onlineCredentials = new SharePointOnlineCredentials(credentials.UserName, credentials.Password);
                    context.Credentials = onlineCredentials;
                    try
                    {
                        context.ExecuteQueryRetry();
                    }
                    catch (IdcrlException iex)
                    {
                        System.Diagnostics.Trace.TraceError("Authentication Exception {0}", iex.Message);
                        return null;
                    }
                    catch (WebException wex)
                    {
                        System.Diagnostics.Trace.TraceError("Authentication Exception {0}", wex.Message);
                        return null;
                    }
                    catch (ClientRequestException)
                    {
                        context.Credentials = new NetworkCredential(credentials.UserName, credentials.Password);
                    }
                    catch (ServerException)
                    {
                        context.Credentials = new NetworkCredential(credentials.UserName, credentials.Password);
                    }
                }
                catch (ArgumentException)
                {
                    // OnPrem?
                    context.Credentials = new NetworkCredential(credentials.UserName, credentials.Password);
                    try
                    {
                        context.ExecuteQueryRetry();
                    }
                    catch (ClientRequestException ex)
                    {
                        throw new Exception("Error establishing a connection", ex);
                    }
                    catch (ServerException ex)
                    {
                        throw new Exception("Error establishing a connection", ex);
                    }
                }

            }
            else
            {
                if (credentials != null)
                {
                    context.Credentials = new NetworkCredential(credentials.UserName, credentials.Password);
                }
            }

            var connectionType = ConnectionType.OnPrem;
            if (url.Host.ToUpperInvariant().EndsWith("SHAREPOINT.COM"))
            {
                connectionType = ConnectionType.O365;
            }

            if (skipAdminCheck == false)
            {
                if (IsTenantAdminSite(context))
                {
                    connectionType = ConnectionType.TenantAdmin;
                }
            }

            return new SPIaCConnection(context, connectionType, minimalHealthScore, retryCount, retryWait, credentials, url.ToString());
        }

        private static bool IsTenantAdminSite(ClientContext clientContext)
        {
            try
            {
                var tenant = new Tenant(clientContext);
                clientContext.ExecuteQueryRetry();
                return true;
            }
            catch (ClientRequestException)
            {
                return false;
            }
            catch (ServerException)
            {
                return false;
            }
        }

    }
}
