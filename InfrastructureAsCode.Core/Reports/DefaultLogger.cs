using System;

namespace InfrastructureAsCode.Core.Reports
{
    public class DefaultLogger : ITraceLogger
    {
        public void LogError(Exception ex, String message, params object[] args)
        {
            System.Diagnostics.Trace.TraceInformation(message, args);
        }

        public void LogInformation(String message, params object[] args)
        {
            System.Diagnostics.Trace.TraceInformation(message, args);
        }
    }
}