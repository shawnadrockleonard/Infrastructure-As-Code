using InfrastructureAsCode.Core.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Powershell.Commands.Base
{
    /// <summary>
    /// Base class for all the Microsoft Graph related cmdlets
    /// </summary>
    public abstract class ExtendedPSCmdlet : PSCmdlet
    {
        /// <summary>
        /// the logger is available
        /// </summary>
        internal bool loggerAvailable { get; private set; }

        /// <summary>
        /// initializer a logger
        /// </summary>
        internal ConfigurationLogger logger { get; private set; }

        /// <summary>
        /// Storage for the cmdlet in the current thread
        /// </summary>
        private string m_cmdLetName { get; set; }
        internal string CmdLetName
        {
            get
            {
                if (string.IsNullOrEmpty(m_cmdLetName))
                {
                    var runningAssembly = Assembly.GetExecutingAssembly();
                    m_cmdLetName = this.GetType().Name;
                }
                return m_cmdLetName;
            }
        }

        /// <summary>
        /// Initializers the logger from the cmdlet
        /// </summary>
        protected virtual void OnBeginInitialize()
        {
        }

        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            var runningDirectory = this.SessionState.Path.CurrentFileSystemLocation;
            var runningAssembly = Assembly.GetExecutingAssembly();
            var runningAssemblyName = runningAssembly.ManifestModule.Name;

            var appConfig = string.Format("{0}\\{1}.config", runningDirectory, runningAssemblyName).Replace("\\", @"\");
            if (System.IO.File.Exists(appConfig))
            {
                LogVerbose("AppSettings file found at {0}", appConfig);
                logger = new ConfigurationLogger(appConfig, true, CmdLetName);
                loggerAvailable = true;
            }

            OnBeginInitialize();
            LogVerbose(">>> Begin {0} at {1}", this.CmdLetName, DateTime.UtcNow);
        }

        public virtual void ExecuteCmdlet()
        { }

        protected override void ProcessRecord()
        {
            ExecuteCmdlet();
        }

        /// <summary>
        /// End Processing cleanup or write logs
        /// </summary>
        protected override void EndProcessing()
        {
            base.EndProcessing();
            LogVerbose("<<< End {0} at {1}", CmdLetName, DateTime.Now);
        }

        /// <summary>
        /// retrieve app setting from app.config
        /// </summary>
        /// <param name="settingName"></param>
        /// <returns></returns>
        protected virtual string GetAppSetting(string settingName)
        {
            if (logger != null)
            {
                return logger.GetAppSetting(settingName);
            }
            return null;
        }

        /// <summary>
        /// retrieve connection string from app.config
        /// </summary>
        /// <param name="settingName"></param>
        /// <returns></returns>
        protected virtual string GetConnectionString(string settingName)
        {
            if (logger != null)
            {
                return logger.GetConnectionSetting(settingName);
            }
            return null;
        }

        /// <summary>
        /// Log: ERROR
        /// </summary>
        /// <param name="ex"></param>
        /// <param name="message"></param>
        /// <param name="args"></param>
        protected virtual void LogError(Exception ex, string message, params object[] args)
        {
            if (loggerAvailable)
            {
                logger.Error(ex, message, args);
            }
            System.Diagnostics.Trace.TraceError(message, args);
            System.Diagnostics.Trace.TraceError("Exception: {0}", ex.Message);
            if (!Stopping)
            {
                WriteError(new ErrorRecord(ex, "HALT", ErrorCategory.FromStdErr, null));
            }
        }

        /// <summary>
        /// Log: DEBUG
        /// </summary>
        /// <param name="message"></param>
        /// <param name="args"></param>
        protected virtual void LogDebugging(string message, params object[] args)
        {
            if (loggerAvailable)
            {
                logger.Debugging(message, args);
            }
            System.Diagnostics.Trace.TraceInformation(message, args);
            if (!Stopping)
            {
                WriteDebug(string.Format(message, args));
            }
        }

        /// <summary>
        /// Writes a warning message to the cmdlet and logs to directory
        /// </summary>
        /// <param name="message"></param>
        /// <param name="args"></param>
        protected virtual void LogWarning(string message, params object[] args)
        {
            if (loggerAvailable)
            {
                logger.Warning(string.Format(message, args));
            }
            System.Diagnostics.Trace.TraceWarning(message, args);
            if (!Stopping)
            {
                WriteWarning(string.Format(message, args));
            }
        }

        /// <summary>
        /// Log: VERBOSE
        /// </summary>
        /// <param name="message"></param>
        /// <param name="args"></param>
        protected virtual void LogVerbose(string message, params object[] args)
        {
            if (loggerAvailable)
            {
                logger.Information(message, args);
            }
            System.Diagnostics.Trace.TraceInformation(message, args);
            if (!Stopping)
            {
                WriteVerbose(string.Format(message, args));
            }
        }
    }
}