using Microsoft.IaC.Powershell.PipeBinds;
using Microsoft.IaC.Core.Extensions;
using Microsoft.SharePoint.Client;
using System;
using System.Management.Automation;

namespace Microsoft.IaC.Powershell.CmdLets
{
    public class SPIaCWebCmdlet : SPIaCCmdlet
    {
        private Microsoft.SharePoint.Client.Web _selectedWeb;


        [Parameter(Mandatory = false, HelpMessage = "The web to apply the command to. Omit this parameter to use the current web.")]
        public WebPipeBind Web = new WebPipeBind();

        internal Microsoft.SharePoint.Client.Web SelectedWeb
        {
            get
            {
                if (_selectedWeb == null)
                {
                    _selectedWeb = GetWeb();
                }
                return _selectedWeb;
            }
        }

        private Microsoft.SharePoint.Client.Web GetWeb()
        {
            var web = ClientContext.Web;

            if (Web.Id != Guid.Empty)
            {
                web = web.GetWebById(Web.Id);
                SPIaCConnection.CurrentConnection.Context = ClientContext.Clone(web.Url);
                web = SPIaCConnection.CurrentConnection.Context.Web;
            }
            else if (!string.IsNullOrEmpty(Web.Url))
            {
                web = web.GetWebByUrl(Web.Url);
                SPIaCConnection.CurrentConnection.Context = ClientContext.Clone(web.Url);
                web = SPIaCConnection.CurrentConnection.Context.Web;
            }
            else if (Web.Web != null)
            {
                web = Web.Web;
                if (!web.IsPropertyAvailable("Url"))
                {
                    ClientContext.Load(web, w => w.Url);
                    ClientContext.ExecuteQueryRetry();
                }
                SPIaCConnection.CurrentConnection.Context = ClientContext.Clone(web.Url);
                web = SPIaCConnection.CurrentConnection.Context.Web;
            }
            else
            {
                if (SPIaCConnection.CurrentConnection.Context.Url != SPIaCConnection.CurrentConnection.Url)
                {
                    SPIaCConnection.CurrentConnection.RestoreCachedContext();
                }
                web = ClientContext.Web;
            }


            return web;
        }

        protected override void EndProcessing()
        {
            base.EndProcessing();
            if (SPIaCConnection.CurrentConnection.Context.Url != SPIaCConnection.CurrentConnection.Url)
            {
                SPIaCConnection.CurrentConnection.RestoreCachedContext();
            }
        }

        protected override void BeginProcessing()
        {
            base.BeginProcessing();
            SPIaCConnection.CurrentConnection.CacheContext();
        }

    }
}