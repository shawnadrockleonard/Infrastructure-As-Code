using InfrastructureAsCode.Core.Reports.o365Graph.AzureAD;
using InfrastructureAsCode.Core.Reports.o365Graph.TenantReport;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Reports.o365Graph
{
    public interface IReportVisitor
    {
        /// <summary>
        /// Will process the report based on the requested <typeparamref name="T"/> class
        ///     NOTE: <paramref name="betaEndPoint"/> true will default to using the JSON format; if an exception is raise it will retry with the CSV format
        ///     NOTE: <paramref name="betaEndPoint"/> false will skip JSON and use the CSV format with v1.0 endpoint
        /// </summary>
        /// <typeparam name="T">GraphAPIReport Models</typeparam>
        /// <param name="reportPeriod"></param>
        /// <param name="reportDate"></param>
        /// <param name="defaultRecordBatch"></param>
        /// <param name="betaEndPoint"></param>
        /// <returns></returns>
        ICollection<T> ProcessReport<T>(ReportUsagePeriodEnum reportPeriod, Nullable<DateTime> reportDate, int defaultRecordBatch = 500, bool betaEndPoint = false) where T : JSONResult;

        /// <summary>
        /// Queries the Graph API and returns the emitted string output
        /// </summary>
        /// <param name="ReportType"></param>
        /// <param name="reportPeriod"></param>
        /// <param name="reportDate"></param>
        /// <param name="defaultRecordBatch">(OPTIONAL) will default to 500</param>
        /// <param name="betaEndPoint">(OPTIONAL) will default to false</param>
        /// <returns></returns>
        string ProcessReport(ReportUsageTypeEnum ReportType, ReportUsagePeriodEnum reportPeriod, Nullable<DateTime> reportDate, int defaultRecordBatch = 500, bool betaEndPoint = false);

    }
}
