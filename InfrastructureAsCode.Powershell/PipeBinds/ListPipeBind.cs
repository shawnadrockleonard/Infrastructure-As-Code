using Microsoft.SharePoint.Client;
using System;
using System.Linq.Expressions;

namespace InfrastructureAsCode.Powershell.PipeBinds
{
    public sealed class ListPipeBind
    {
        private readonly List _list;
        private readonly Guid _id;
        private readonly string _name;

        public ListPipeBind()
        {
            _list = null;
            _id = Guid.Empty;
            _name = string.Empty;
        }

        public ListPipeBind(List list)
        {
            _list = list;
        }

        public ListPipeBind(Guid guid)
        {
            _id = guid;
        }

        public ListPipeBind(string id)
        {
            if (!Guid.TryParse(id, out _id))
            {
                _name = id;
            }
        }

        public Guid Id
        {
            get { return _id; }
        }

        public List List
        {
            get
            {
                return _list;
            }
        }

        public string Title
        {
            get { return _name; }
        }

        internal List GetList(Web web, params Expression<Func<List, object>>[] expressions)
        {
            List list = null;
            if (List != null)
            {
                list = List;
            }
            else if (Id != Guid.Empty)
            {
                list = web.Lists.GetById(Id);
            }
            else if (!string.IsNullOrEmpty(Title))
            {
                list = web.GetListByTitle(Title, expressions);
                if (list == null)
                {
                    list = web.GetListByUrl(Title, expressions);
                }
            }

            // if no expressions are supplied then use default settings
            if (list != null && expressions.Length <= 0)
            {
                web.Context.Load(list, l => l.Id, l => l.BaseTemplate, l => l.OnQuickLaunch, l => l.DefaultViewUrl, l => l.Title, l => l.Hidden, l => l.ContentTypesEnabled, l => l.RootFolder.ServerRelativeUrl);
                web.Context.ExecuteQueryRetry();
            }
            return list;
        }

        public override string ToString()
        {
            if (List != null)
            {
                return List.Title;
            }
            else if (Id != Guid.Empty)
            {
                return Id.ToString("B");
            }
            else if (!string.IsNullOrEmpty(Title))
            {
                return Title;
            }
            return base.ToString();
        }
    }
}
