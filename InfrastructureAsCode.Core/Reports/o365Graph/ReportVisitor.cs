using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Reports.o365Graph
{
    public abstract class ReportVisitor
    {
        public ITraceLogger Logger { get; set; }

        /// <summary>
        /// Generic processor for the Stream Reader
        /// </summary>
        /// <param name="responseReader"></param>
        public abstract void ProcessReport(StreamReader responseReader);
    }
}
