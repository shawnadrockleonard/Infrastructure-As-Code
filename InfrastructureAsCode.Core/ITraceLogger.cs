using System;

namespace InfrastructureAsCode.Core
{
    public interface ITraceLogger
    {
        /// <summary>
        /// Log information
        /// </summary>
        /// <param name="format">A composite format string.</param>
        /// <param name="args">An object array that contains zero or more objects to format.</param>
        void LogInformation(String format, params object[] args);

        /// <summary>
        /// Log warning
        /// </summary>
        /// <param name="format">A composite format string.</param>
        /// <param name="args">An object array that contains zero or more objects to format.</param>
        void LogWarning(String format, params object[] args);

        /// <summary>
        /// Log exception and 
        /// </summary>
        /// <param name="ex">Exception</param>
        /// <param name="format">A composite format string.</param>
        /// <param name="args">An object array that contains zero or more objects to format.</param>
        void LogError(Exception ex, String format, params object[] args);
    }
}