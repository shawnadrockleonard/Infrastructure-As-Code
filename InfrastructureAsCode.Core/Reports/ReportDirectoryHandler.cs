using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Reports
{
    /// <summary>
    /// Provides consistent pattern for configuring log, xlsx, and template xlsx files
    /// </summary>
    public class ReportDirectoryHandler
    {
        public ReportDirectoryHandler()
        {
            _totalRows = 0;
        }


        public ReportDirectoryHandler(string rootFolder, string fileName, DateTime dateTime, ITraceLogger traceLogger)
            : this()
        {
            _directoryPath = new DirectoryInfo(rootFolder);
            if (!_directoryPath.Exists)
            {
                _directoryPath.Create();
            }

            var steamFolder = "steams";
            _steamPath = new DirectoryInfo(Path.Combine(_directoryPath.FullName, steamFolder));
            if (!_steamPath.Exists)
            {
                _steamPath = _directoryPath.CreateSubdirectory(steamFolder, _directoryPath.GetAccessControl());
            }

            var logFolder = "logs";
            _logPath = new DirectoryInfo(Path.Combine(_directoryPath.FullName, logFolder));
            if (!_logPath.Exists)
            {
                _logPath = _directoryPath.CreateSubdirectory(logFolder, _directoryPath.GetAccessControl());
            }

            TraceLogger = traceLogger;

            _fileName = fileName;
            _dateTime = dateTime;
        }

        private ITraceLogger TraceLogger { get; set; }

        public System.IO.DirectoryInfo _directoryPath { get; private set; }

        public System.IO.DirectoryInfo _steamPath { get; private set; }

        public System.IO.DirectoryInfo _logPath { get; private set; }

        public DateTime _dateTime { get; private set; }


        internal int _totalRows { get; private set; }
        public int TotalRows
        {
            get
            {
                return _totalRows;
            }
        }

        public string _fileName { get; set; }

        public string _fileWithDateName
        {
            get
            {
                return $"{_fileName}_{_dateTime.ToString("yyyy_MM_dd")}";
            }
        }

        public string _logFilePath
        {
            get
            {
                return System.IO.Path.Combine(_logPath.FullName, _fileWithDateName + ".txt");
            }
        }

        public string _logCSVFilePath
        {
            get
            {
                return System.IO.Path.Combine(_logPath.FullName, _fileWithDateName + ".csv");
            }
        }


        /// <summary>
        /// Erases the log file
        /// </summary>
        public void ResetLogFile()
        {
            try
            {
                if (System.IO.File.Exists(_logFilePath))
                {
                    TraceLogger.LogInformation("Deleting LOG file {0}", _logFilePath);
                    System.IO.File.Delete(_logFilePath);
                }
            }
            catch (Exception e)
            {
                TraceLogger.LogError(e, "resetLogFile failed with message {0}", e.Message);
            }
        }

        /// <summary>
        /// Erases the log file
        /// </summary>
        public void ResetCSVFile()
        {
            try
            {
                if (System.IO.File.Exists(_logCSVFilePath))
                {
                    TraceLogger.LogInformation("Deleting CSV file {0}", _logCSVFilePath);
                    System.IO.File.Delete(_logFilePath);
                }
            }
            catch (Exception e)
            {
                TraceLogger.LogError(e, "resetCSVFile failed with message {0}", e.Message);
            }
        }

        public void RemoveSteamFiles()
        {
            // remove steam progress files
            var steamfiles = System.IO.Directory.EnumerateFiles(_steamPath.FullName);
            foreach (var steamfile in steamfiles)
            {
                System.IO.File.Delete(steamfile);
            }
        }

        /// <summary>
        /// Appends a line to the log file path
        /// </summary>
        /// <param name="_line"></param>
        public void WriteToLogFile(string _line)
        {
            if (!string.IsNullOrEmpty(_line))
            {
                try
                {
                    _totalRows++;
                    System.IO.File.AppendAllLines(_logFilePath, new[] { _line });
                }
                catch (Exception e)
                {
                    TraceLogger.LogError(e, "WriteToLogFile failed with message {0}", e.Message);
                }
            }
        }

        /// <summary>
        /// Appends a line to the CSV file path
        /// </summary>
        /// <param name="_line"></param>
        public void WriteToCSVFile(string _line)
        {
            if (!string.IsNullOrEmpty(_line))
            {
                try
                {
                    _totalRows++;
                    System.IO.File.AppendAllLines(_logCSVFilePath, new[] { _line });
                }
                catch (Exception e)
                {
                    TraceLogger.LogError(e, "WriteToCSVFile failed with message {0}", e.Message);
                }
            }
        }
    }

}
