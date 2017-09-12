using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Reports.o365Graph
{
    public class DefaultReportVisitor : ReportVisitor
    {
        public DefaultReportVisitor() { }

        public DefaultReportVisitor(ITraceLogger logger)
        {
            Logger = logger;
        }

        public override void ProcessReport(StreamReader responseReader)
        {
            var response = responseReader.ReadToEnd();
            Logger.LogInformation("WebResponse:{0}", response);
        }
    }
}
