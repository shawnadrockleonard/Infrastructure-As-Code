using Microsoft.SharePoint.Client;

namespace Microsoft.IaC.Powershell.CmdLets
{
    public interface IIaCCmdlet
    {
        ClientContext ClientContext { get; }

        void ExecuteCmdlet();
    }
}