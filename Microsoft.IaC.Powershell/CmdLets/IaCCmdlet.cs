using IaC.Core.Utilities;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Utilities;
using System;
using System.Management.Automation;
using System.Reflection;
using System.Threading;
using System.Xml.Linq;
using Resources = IaC.Core.Properties.Resources;

namespace IaC.Powershell.CmdLets
{
    public abstract class IaCCmdlet : PSCmdlet, IIaCCmdlet
    {
        /// <summary>
        /// If True then only write verbose statements to the log and do not perform any action
        /// </summary>
        [Parameter(Mandatory = false)]
        public SwitchParameter DoNothing { get; set; }

        public ClientContext ClientContext
        {
            get { return SPIaCConnection.CurrentConnection.Context; }
        }

        /// <summary>
        /// The base URI for the SP Site or Tenant
        /// </summary>
        internal string BaseUri { get; private set; }

        /// <summary>
        /// the logger is available
        /// </summary>
        internal bool loggerAvailable { get; private set; }

        /// <summary>
        /// initializer a logger
        /// </summary>
        internal ConfigurationLogger logger { get; private set; }

        /// <summary>
        /// The application setting config manager
        /// </summary>
        internal ConfigurationReader appSettings { get; private set; }

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
        /// Processed before the Execute
        /// </summary>
        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            if (SPIaCConnection.CurrentConnection == null)
            {
                throw new InvalidOperationException(Resources.NoConnection);
            }

            if (ClientContext == null)
            {
                throw new InvalidOperationException(Resources.NoConnection);
            }

            Uri uri = new Uri(this.ClientContext.Url);
            var urlParts = uri.Authority.Split(new[] { '.' });
            BaseUri = string.Format("https://{0}.{1}.{2}", urlParts[0], urlParts[1], urlParts[2]);

            LogVerbose(">>> Begin {0} at {1} on URL:[{2}] [DoNothing:{3}]", this.CmdLetName, DateTime.Now, this.ClientContext.Url, this.DoNothing);
        }

        /// <summary>
        /// Initializers the logger from the cmdlet
        /// </summary>
        protected virtual void PreInitialize()
        {
            var runningDirectory = this.SessionState.Path.CurrentFileSystemLocation;
            var runningAssembly = Assembly.GetExecutingAssembly();

            var appConfig = string.Format("{0}\\{1}.config", runningDirectory, runningAssembly.ManifestModule.Name).Replace("\\", @"\");
            if (System.IO.File.Exists(appConfig))
            {
                LogVerbose("AppSettings file found at {0}", appConfig);
                appSettings = new ConfigurationReader(appConfig);
            }

            var logConfig = string.Format("{0}\\{1}.config", runningDirectory, CmdLetName).Replace("\\", @"\");
            if (System.IO.File.Exists(logConfig))
            {
                LogVerbose("Configuration file found at {0}", logConfig);
                loggerAvailable = true;
                logger = new ConfigurationLogger(logConfig);
            }
            else
            {
                var samplelogConfig = string.Format("{0}\\samplelog.config", runningDirectory).Replace("\\", @"\");
                if (System.IO.File.Exists(samplelogConfig))
                {
                    System.IO.File.Copy(samplelogConfig, logConfig, false);
                    XDocument xmlConfig = null;
                    using (var sr = new System.IO.StreamReader(samplelogConfig))
                    {
                        // Read the stream to a string, and write the string to the console.
                        var reader = new System.Xml.XmlTextReader(sr);
                        xmlConfig = XDocument.Load(reader);
                        var xmlElement = xmlConfig.Root.Element("log4net").Element("appender").Element("file");
                        var xmlElementValue = xmlElement.Attribute("value").Value.Replace("samplelogfolder", CmdLetName);
                        xmlElement.SetAttributeValue("value", xmlElementValue);
                    }
                    xmlConfig.Save(logConfig, SaveOptions.DisableFormatting);
                    LogVerbose("Configuration file written to {0}", logConfig);
                    loggerAvailable = true;
                    logger = new ConfigurationLogger(logConfig);
                }
            }

        }

        /// <summary>
        /// Execute custom cmdlet code
        /// </summary>
        public virtual void ExecuteCmdlet()
        {
        }

        /// <summary>
        /// Process SPO HealthCheck and validation context
        /// </summary>
        protected override void ProcessRecord()
        {
            try
            {
                if (SPIaCConnection.CurrentConnection.MinimalHealthScore != -1)
                {
                    int healthScore = Utility.GetHealthScore(SPIaCConnection.CurrentConnection.Url);
                    if (healthScore <= SPIaCConnection.CurrentConnection.MinimalHealthScore)
                    {
                        PreInitialize();
                        ExecuteCmdlet();
                    }
                    else
                    {
                        if (SPIaCConnection.CurrentConnection.RetryCount != -1)
                        {
                            int retry = 1;
                            while (retry <= SPIaCConnection.CurrentConnection.RetryCount)
                            {
                                WriteWarning(string.Format(Resources.Retry0ServerNotHealthyWaiting1seconds, retry, SPIaCConnection.CurrentConnection.RetryWait, healthScore));
                                Thread.Sleep(SPIaCConnection.CurrentConnection.RetryWait * 1000);
                                healthScore = Utility.GetHealthScore(SPIaCConnection.CurrentConnection.Url);
                                if (healthScore <= SPIaCConnection.CurrentConnection.MinimalHealthScore)
                                {
                                    PreInitialize();
                                    ExecuteCmdlet();
                                    break;
                                }
                                retry++;
                            }
                            if (retry > SPIaCConnection.CurrentConnection.RetryCount)
                            {
                                WriteError(new ErrorRecord(new Exception(Resources.HealthScoreNotSufficient), "HALT", ErrorCategory.OpenError, null));
                            }
                        }
                        else
                        {
                            WriteError(new ErrorRecord(new Exception(Resources.HealthScoreNotSufficient), "HALT", ErrorCategory.OpenError, null));
                        }
                    }
                }
                else
                {
                    PreInitialize();
                    ExecuteCmdlet();
                }
            }
            catch (Exception ex)
            {
                SPIaCConnection.CurrentConnection.RestoreCachedContext();
                WriteError(new ErrorRecord(ex, "EXCEPTION", ErrorCategory.WriteError, null));
            }
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
        /// internal member to hold the current user
        /// </summary>
        private string _currentUserInProcess = string.Empty;

        /// <summary>
        /// this should be valid based on pre authentication checks
        /// </summary>
        protected virtual string CurrentUserName
        {
            get
            {
                if (string.IsNullOrEmpty(_currentUserInProcess))
                {
                    try
                    {
                        _currentUserInProcess = SPIaCConnection.CurrentConnection.GetActiveUsername();
                    }
                    catch (Exception ex)
                    {
                        LogError(ex, "Failed to retrieve the current context credential for the Cached Entity");
                    }
                }
                return _currentUserInProcess;
            }
        }

        /// <summary>
        /// retrieve app setting from app.config
        /// </summary>
        /// <param name="settingName"></param>
        /// <returns></returns>
        protected virtual string GetAppSetting(string settingName)
        {
            if (appSettings != null)
            {
                return appSettings.GetAppSetting(settingName);
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
            if (appSettings != null)
            {
                return appSettings.GetConnectionSetting(settingName);
            }
            return null;
        }

        /// <summary>
        /// Log: ERROR
        /// </summary>
        /// <param name="ex"></param>
        /// <param name="category"></param>
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
            WriteError(new ErrorRecord(ex, "HALT", ErrorCategory.FromStdErr, null));
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
            WriteDebug(string.Format(message, args));
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
            WriteWarning(string.Format(message, args));
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
            WriteVerbose(string.Format(message, args));
        }
    }
}
