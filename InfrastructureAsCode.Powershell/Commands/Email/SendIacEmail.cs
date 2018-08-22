using InfrastructureAsCode.Powershell.Commands.Base;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Extensions;

namespace InfrastructureAsCode.Powershell.Commands.Email
{
    [Cmdlet(VerbsExtended.Send, "IacEmail", SupportsShouldProcess = true)]
    public class SendIacEmail : IaCCmdlet
    {
        #region Parameters

        [Parameter(Mandatory = true)]
        public string[] Emails { get; set; }

        [Parameter(Mandatory = false)]
        public string Subject { get; set; }

        [Parameter(Mandatory = false)]
        public string Body { get; set; }

        #endregion

        public override void ExecuteCmdlet()
        {
            base.ExecuteCmdlet();


            var properties = new Microsoft.SharePoint.Client.Utilities.EmailProperties
            {
                To = Emails,
                Subject = Subject,
                Body = Body
            };

            Microsoft.SharePoint.Client.Utilities.Utility.SendEmail(this.ClientContext, properties);
            ClientContext.ExecuteQueryRetry();

        }
    }
}
