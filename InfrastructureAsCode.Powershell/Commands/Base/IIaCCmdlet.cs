using Microsoft.SharePoint.Client;

namespace InfrastructureAsCode.Powershell.Commands.Base
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