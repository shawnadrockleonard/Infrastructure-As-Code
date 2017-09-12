using InfrastructureAsCode.Powershell.CmdLets;
using System.Management.Automation;


namespace InfrastructureAsCode.Powershell.Commands
{
    /// <summary>
    /// Returns a Client Side Object Model context
    /// </summary>
    [Cmdlet(VerbsCommon.Get, "SPIaCContext")]
    public class GetSPIaCContext : PSCmdlet
    {
        protected override void ProcessRecord()
        {
            WriteObject(SPIaCConnection.CurrentConnection.Context);
        }
    }
}
