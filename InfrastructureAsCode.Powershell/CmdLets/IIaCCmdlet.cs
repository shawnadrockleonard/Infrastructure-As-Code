using Microsoft.SharePoint.Client;

namespace InfrastructureAsCode.Powershell.CmdLets
{
    /// <summary>
    /// Interface for every command to implement
    /// </summary>
    public interface IIaCCmdlet
    {
        ClientContext ClientContext { get; }

        void ExecuteCmdlet();
    }
}