using System;

namespace InfrastructureAsCode.Core.Reports.o365rwsclient
{
    public class DefaultReportVisitor : ReportVisitor
    {
        /// <summary>
        ///
        /// </summary>
        /// <param name="report"></param>
        public override void VisitReport(ReportObject report)
        {
            System.Diagnostics.Trace.TraceInformation(report.ConvertToXml());
        }

        /// <summary>
        ///
        /// </summary>
        public override void VisitBatchReport()
        {
            foreach (ReportObject report in this.reportObjectList)
            {
                System.Diagnostics.Trace.TraceInformation(report.ConvertToXml());
            }
        }
    }
}