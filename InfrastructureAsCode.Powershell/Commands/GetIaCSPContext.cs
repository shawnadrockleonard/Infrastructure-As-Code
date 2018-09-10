using InfrastructureAsCode.Powershell.Commands.Base;
using System.Management.Automation;


namespace InfrastructureAsCode.Powershell.Commands
{
    /// <summary>
    /// Returns a Client Side Object Model context
    /// </summary>
    [Cmdlet(VerbsCommon.Get, "IaCSPContext")]
    public class GetIaCSPContext : PSCmdlet
    {
        protected override void ProcessRecord()
        {
            WriteObject(SPIaCConnection.CurrentConnection.Context);
        }
    }
}
