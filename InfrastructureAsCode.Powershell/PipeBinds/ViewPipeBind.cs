using Microsoft.SharePoint.Client;
using System;

namespace InfrastructureAsCode.Powershell.PipeBinds
{
    public class ViewPipeBind
    {
        private readonly View _view;
        private readonly Guid _id;
        private readonly string _name;

        public ViewPipeBind()
        {
            _view = null;
            _id = Guid.Empty;
            _name = string.Empty;
        }

        public ViewPipeBind(View view)
        {
            _view = view;
        }

        public ViewPipeBind(Guid guid)
        {
            _id = guid;
        }

        public ViewPipeBind(string id)
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

        public View View
        {
            get
            {
                return _view;
            }
        }

        public string Title
        {
            get { return _name; }
        }

        internal View GetView(List list)
        {
            View view = null;
            if (View != null)
            {
                view = View;
            }
            else if (Id != Guid.Empty)
            {
                view = list.GetViewById(Id);
            }
            else if (!string.IsNullOrEmpty(Title))
            {
                view = list.GetViewByName(Title);
            }

            if (view != null)
            {
                list.Context.Load(view, 
                    tv => tv.Id,
                    tv => tv.Title,
                    tv => tv.ServerRelativeUrl,
                    tv => tv.DefaultView,
                    tv => tv.HtmlSchemaXml, 
                    tv => tv.RowLimit,
                    tv => tv.Toolbar,
                    tv => tv.JSLink, 
                    tv => tv.ViewFields,
                    tv => tv.ViewQuery,
                    tv => tv.Aggregations,
                    tv => tv.AggregationsStatus,
                    tv => tv.Hidden,
                    tv => tv.Method,
                    tv => tv.PersonalView,
                    tv => tv.ReadOnlyView,
                    tv => tv.ViewType);
                list.Context.ExecuteQueryRetry();
            }

            return view;
        }
    }
}
