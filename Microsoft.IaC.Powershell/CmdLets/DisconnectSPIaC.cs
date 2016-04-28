﻿using System;
using System.Management.Automation;
using Resources = Microsoft.IaC.Powershell.Properties.Resources;

namespace Microsoft.IaC.Powershell.CmdLets
{
    /*
        Examples:
        Disconnect-SPOnline
    */
    [Cmdlet("Disconnect", "SPIaC")]
    [CmdletHelp("Disconnects the context", Category = "Base Cmdlets")]
    public class DisconnectSPIaC : PSCmdlet
    {
        protected override void ProcessRecord()
        {
            if (!DisconnectCurrentService())
                throw new InvalidOperationException(Resources.NoConnectionToDisconnect);
        }

        internal static bool DisconnectCurrentService()
        {
            if (SPIaCConnection.CurrentConnection == null)
                return false;
            SPIaCConnection.CurrentConnection = null;
            return true;
        }
    }
}
