using System.Management.Automation;
using InfrastructureAsCode.Powershell.Extensions;
using OfficeDevPnP.Core.Utilities;

namespace InfrastructureAsCode.Powershell.PipeBinds
{
    public sealed class CredentialPipeBind
    {
        private readonly PSCredential _pscredential;
        private readonly string _storedcredential;

        public CredentialPipeBind(PSCredential pscredential)
        {
            _pscredential = pscredential;
        }

        public CredentialPipeBind(string id)
        {
            _storedcredential = id;
        }

        public PSCredential Credential
        {
            get
            {
                if (_pscredential != null)
                {
                    return _pscredential;
                }
                else if (_storedcredential != null)
                {
                    var credsPtr = CredentialManager.GetCredential(_storedcredential);
                    return credsPtr.GetPSCredentials();
                }
                else
                {
                    return null;
                }
            }
        }
    }
}
