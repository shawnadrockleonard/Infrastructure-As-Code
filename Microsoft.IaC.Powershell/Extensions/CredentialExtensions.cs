using System;
using System.Management.Automation;
using System.Runtime.InteropServices;
using System.Security;
using Microsoft.Win32.SafeHandles;
using FILETIME = System.Runtime.InteropServices.ComTypes.FILETIME;
using Microsoft.SharePoint.Client;
using System.Net;

namespace Microsoft.IaC.Powershell.Extensions
{
    internal static class CredentialExtensions
    {

        public static PSCredential GetPSCredentials(this SharePointOnlineCredentials credentials)
        {
            var psCredential = credentials.GetPSCredentials();
            return psCredential;
        }

        public static PSCredential GetPSCredentials(this NetworkCredential credentials)
        {
            var psCredential = new PSCredential(credentials.UserName, credentials.SecurePassword);
            return psCredential;
        }

    }
}
