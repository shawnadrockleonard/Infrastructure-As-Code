using InfrastructureAsCode.Powershell.PipeBinds;
using InfrastructureAsCode.Core.Extensions;
using Microsoft.SharePoint.Client;
using System;
using System.Management.Automation;

namespace InfrastructureAsCode.Powershell.CmdLets
{
    /// <summary>
    /// Represents a SPWeb instance
    /// </summary>
    public class IaCWebCmdlet : IaCCmdlet
    {
        private Web _selectedWeb;


        [Parameter(Mandatory = false, HelpMessage = "The web to apply the command to. Omit this parameter to use the current web.")]
        public WebPipeBind Web = new WebPipeBind();

        internal Web SelectedWeb
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

        private Web GetWeb()
        {
            Web web = ClientContext.Web;

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