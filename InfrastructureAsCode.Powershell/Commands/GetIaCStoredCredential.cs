using System.Management.Automation;
using InfrastructureAsCode.Core.Enums;
using InfrastructureAsCode.Powershell.Extensions;
using InfrastructureAsCode.Powershell.CmdLets;

namespace InfrastructureAsCode.Powershell.Commands
{
    /*
    Example:
    Returns the credential associated with the specified identifier
    Get-IaCStoredCredential -Name O365
    */
    [Cmdlet("Get", "IaCStoredCredential")]
    [CmdletHelp("Returns a stored credential from the Windows Credential Manager", Category = "Base Cmdlets")]
    public class GetIaCStoredCredential : PSCmdlet
    {
        [Parameter(Mandatory = true, HelpMessage = "The credential to retrieve.")]
        public string Name;

        [Parameter(Mandatory = false, HelpMessage = "The object type of the credential to return from the Credential Manager. Possible valus are 'O365', 'OnPrem' or 'PSCredential'")]
        public CredentialType Type = CredentialType.O365;

        protected override void ProcessRecord()
        {
            switch (Type)
            {
                case CredentialType.O365:
                    {
                        WriteObject(OfficeDevPnP.Core.Utilities.CredentialManager.GetSharePointOnlineCredential(Name));
                        break;
                    }
                case CredentialType.OnPrem:
                    {
                        var item = CredentialManager.GetCredential(Name);
                        WriteObject(item);
                        break;
                    }
                case CredentialType.PSCredential:
                    {
                        var item = CredentialManager.GetCredential(Name);
                        WriteObject(item);
                        break;
                    }
            }
        }
    }
}
