using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Powershell.CmdLets
{
    /// Provides an extended set of "unsupported" powershell cmdlets
    /// </summary>
    public static class VerbsExtended
    {
        public const string Connect = "Connect";
        public const string Sync = "Sync";
        public const string Import = "Import";
        public const string Send = "Send";
        public const string Scan = "Scan";
        public const string Report = "Report";
        public const string Disconnect = "Disconnect";
    }
}
