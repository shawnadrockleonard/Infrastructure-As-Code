using InfrastructureAsCode.Core.Reports.o365Graph.TenantReport;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Reports.o365Graph
{
    public class DefaultReportVisitor : ReportVisitor, IDisposable
    {
        #region Properties

        public bool _disposed { get; set; }

        #endregion

        public DefaultReportVisitor() { }

        public DefaultReportVisitor(ITraceLogger logger)
        {
            Logger = logger;
        }


        public override void ProcessReport(ReportingStream responseReader, QueryFilter serviceQuery)
        {
            var response = responseReader.RetrieveData(serviceQuery);
            Logger.LogInformation("WebResponse:{0}", response);
        }


        protected virtual void Dispose(bool disposing)
        {
            if (disposing
                && !_disposed)
            {
            }
            _disposed = true;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
    }
}
