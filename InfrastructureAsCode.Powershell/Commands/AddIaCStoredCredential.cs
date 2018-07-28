using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using InfrastructureAsCode.Powershell.CmdLets;

namespace InfrastructureAsCode.Powershell.Commands
{
     [Cmdlet(VerbsCommon.Add, "IaCStoredCredential")]
    public class AddIaCStoredCredential : PSCmdlet
    {
        [Parameter(Mandatory = true, HelpMessage = "The credential to set")]
        public string Name;

        [Parameter(Mandatory = true)]
        public string Username;

        [Parameter(Mandatory = false, HelpMessage = @"If not specified you will be prompted to enter your password. If you want to specify this value use ConvertTo-SecureString -String 'YourPassword' -AsPlainText -Force")]
        public SecureString Password;

#if NETSTANDARD2_0
        [Parameter(Mandatory = false)]
        public SwitchParameter Overwrite;
#endif

        protected override void ProcessRecord()
        {
            if (Password == null || Password.Length == 0)
            {
                Host.UI.Write("Enter password: ");
                Password = Host.UI.ReadLineAsSecureString();
            }

#if !NETSTANDARD2_0
            CredentialManager.AddCredential(Name, Username, Password);
#else
            CredentialManager.AddCredential(Name, Username, Password, Overwrite.ToBool());
#endif
        }
    }
}