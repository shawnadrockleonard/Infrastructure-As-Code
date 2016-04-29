using Microsoft.SharePoint.Client;

namespace IaC.Powershell.CmdLets
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