using log4net;
using log4net.Config;
using log4net.Repository.Hierarchy;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.IaC.Core.Utilities
{
    /// <summary>
    /// The Logger class provides wrapper methods to the LogManager.
    /// </summary>
    public class ConfigurationLogger
    {
        /// <summary>
        /// The log
        /// </summary>
        internal static readonly ILog log = LogManager.GetLogger(typeof(Logger));

        /// <summary>
        /// Initializes the <see cref="Logger"/> class.
        /// </summary>
        public ConfigurationLogger()
        {
            XmlConfigurator.Configure();
        }

        /// <summary>
        /// Initializes the <see cref="Logger"/> class.
        /// </summary>
        /// <param name="options">option file for the log config override</param>
        public ConfigurationLogger(string options)
        {
            if(!System.IO.File.Exists(options))
            {
                throw new System.IO.FileNotFoundException(string.Format("File {0} could not be found.", options));
            }
            XmlConfigurator.Configure(new System.IO.FileInfo(options));
        }

        /// <summary>
        /// Logs the specified message as a debug statement
        /// </summary>
        /// <param name="fmt">The message to be logged</param>
        /// <param name="vars"></param>
        public void Debugging(string fmt, params object[] vars)
        {
            var message = string.Format(fmt, vars);
            log.Debug(message);
        }

        /// <summary>
        /// Logs the specified formatted message string with arguments
        /// </summary>
        /// <param name="fmt"></param>
        /// <param name="vars"></param>
        public void Information(string fmt, params object[] vars)
        {
            string message;
            if (vars != null && vars.Length > 0)
            {
                message = string.Format(fmt, vars);
            }
            else
            {
                message = fmt;
            }

            log.Info(message);
        }

        /// <summary>
        /// Logs the specified message as a warning statement
        /// </summary>
        /// <param name="message">The message to be logged</param>
        /// <param name="ex">The exception to be included in the log</param>
        public void Warning(string message, Exception ex = null)
        {
            log.Warn(message, ex);
        }

        /// <summary>
        /// Logs the specified message as a warning statement
        /// </summary>
        /// <param name="fmt">The message to be logged</param>
        /// <param name="vars">collection of values to be injected in string format</param>
        public void Warning(string fmt, params object[] vars)
        {
            Warning(string.Format(fmt, vars));
        }

        /// <summary>
        /// Logs the specified message as an error
        /// </summary>
        /// <param name="message">The message to be logged</param>
        public void Error(string message)
        {
            log.Error(message);
        }

        /// <summary>
        /// Logs the specified message as an error
        /// </summary>
        /// <param name="ex">The exception to be logged</param>
        /// <param name="message">The message to be logged</param>
        public void Error(Exception ex, string message)
        {
            log.Error(message, ex);
        }

        /// <summary>
        /// Logs the exception with the specified message
        /// </summary>
        /// <param name="ex"></param>
        /// <param name="fmt"></param>
        /// <param name="vars"></param>
        public void Error(Exception ex, string fmt, params object[] vars)
        {
            log.Error(string.Format(fmt, vars), ex);
        }

        /// <summary>
        /// Simple exception formatting: for a more comprehensive version see 
        ///     http://code.msdn.microsoft.com/windowsazure/Fix-It-app-for-Building-cdd80df4
        /// </summary>
        /// <param name="exception"></param>
        /// <param name="fmt"></param>
        /// <param name="vars"></param>
        /// <returns></returns>
        private string FormatExceptionMessage(Exception exception, string fmt, object[] vars)
        {
            var sb = new StringBuilder();
            sb.Append(string.Format(fmt, vars));
            sb.Append(" Exception: ");
            sb.Append(exception.ToString());
            return sb.ToString();
        }
    }
}
