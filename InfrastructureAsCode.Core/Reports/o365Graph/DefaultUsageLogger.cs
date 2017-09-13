using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Reports.o365Graph
{
    public class DefaultUsageLogger : ITraceLogger
    {
        private Action<Exception, string, object[]> actionError;
        private Action<string, object[]> actionWarning;
        private Action<string, object[]> actionInformation;

        public DefaultUsageLogger()
        {
            actionError = (Exception ex, string arg1, object[] arg2) =>
            {
                System.Diagnostics.Trace.TraceError(arg1, arg2);
            };
            actionWarning = (string arg1, object[] arg2) =>
            {
                System.Diagnostics.Trace.TraceWarning(arg1, arg2);
            };
            actionInformation = (string arg1, object[] arg2) =>
            {
                System.Diagnostics.Trace.TraceInformation(arg1, arg2);
            };

        }

        public DefaultUsageLogger(
            Action<string, object[]> _actionInformation,
            Action<string, object[]> _actionWarning,
            Action<Exception, string, object[]> _actionError)
        {
            actionError = _actionError;
            actionWarning = _actionWarning;
            actionInformation = _actionInformation;
        }

        public void LogError(Exception ex, string format, params object[] args)
        {
            actionError(ex, format, args);
        }

        public void LogWarning(string format, params object[] args)
        {
            actionWarning(format, args);
        }

        public void LogInformation(string format, params object[] args)
        {
            actionInformation(format, args);
        }
    }
}
