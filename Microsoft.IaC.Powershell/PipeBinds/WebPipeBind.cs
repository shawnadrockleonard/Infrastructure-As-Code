using Microsoft.SharePoint.Client;
using System;

namespace IaC.Powershell.PipeBinds
{
    public sealed class WebPipeBind
    {
        private readonly Guid _id;
        private readonly string _url;
        private readonly Microsoft.SharePoint.Client.Web _web;

        public WebPipeBind()
        {
            _id = Guid.Empty;
            _url = string.Empty;
            _web = null;
        }

        public WebPipeBind(Guid guid)
        {
            _id = guid;
        }

        public WebPipeBind(string id)
        {
            if (!Guid.TryParse(id, out _id))
            {
                _url = id;
            }
        }

        public WebPipeBind(Microsoft.SharePoint.Client.Web web)
        {
            _web = web;
        }

        public Guid Id
        {
            get { return _id; }
        }

        public string Url
        {
            get { return _url; }
        }

        public Microsoft.SharePoint.Client.Web Web
        {
            get
            {
                return _web;
            }
        }


    }
}
