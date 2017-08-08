using Microsoft.SharePoint.Client.WorkflowServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Powershell.PipeBinds
{
    public sealed class WorkflowInstancePipeBind
    {
        private readonly WorkflowInstance _instance;
        private readonly Guid _id;

        public WorkflowInstancePipeBind()
        {
            _instance = null;
            _id = Guid.Empty;
        }

        public WorkflowInstancePipeBind(WorkflowInstance instance)
        {
            _instance = instance;
        }

        public WorkflowInstancePipeBind(Guid guid)
        {
            _id = guid;
        }

        public WorkflowInstancePipeBind(string id)
        {
            _id = Guid.Parse(id);
        }

        public Guid Id
        {
            get { return _id; }
        }

        public WorkflowInstance Instance
        {
            get
            {
                return _instance;
            }
        }
    }
}
