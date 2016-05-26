using InfrastructureAsCode.Powershell.CmdLets;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Powershell.HowToExtend
{
    [Cmdlet(VerbsCommon.Select, "IaCSampleQuery")]
    [CmdletHelp("Demonstrates extending the IaC project", Category = "Query")]
    public class SelectIaCSampleQuery : IaCCmdlet
    {

        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();

            LogVerbose("Successfull ran cmdlet at {0}", DateTime.Now);
        }
    }
}
