using System;

namespace InfrastructureAsCode.Core.Reports.o365rwsclient
{
    /// <summary>
    /// Legacy reporting services are a stop gap until Graph API endpoints are fully functional
    /// </summary>
    [Obsolete("This context class will be officially deprecated on October 1st, 2017")]
    public class ReportingContext
    {
        #region Privates

        private static string defaultServiceEndpointUrl = "https://reports.office365.com/ecp/reportingwebservice/reporting.svc";

        private ITraceLogger logger;

        private ReportVisitor visitor;

        #endregion Privates

        #region Properties

        public string WebServiceUrl
        {
            get;
            set;
        }

        public string UserName
        {
            get;
            set;
        }

        public string Password
        {
            get;
            set;
        }

        public DateTime FromDateTime
        {
            get;
            set;
        }

        public DateTime ToDateTime
        {
            get;
            set;
        }

        /// <summary>
        ///
        /// </summary>
        public string DataFilter
        {
            get;
            set;
        }

        public ITraceLogger TraceLogger
        {
            get
            {
                return this.logger;
            }
        }

        public ReportVisitor ReportVisitor
        {
            get
            {
                return this.visitor;
            }
        }

        #endregion Properties

        #region Constructors

        public ReportingContext()
            : this(defaultServiceEndpointUrl)
        {
        }

        public ReportingContext(string url)
        {
            this.WebServiceUrl = url;
            this.FromDateTime = DateTime.MinValue;
            this.ToDateTime = DateTime.MinValue;
            this.DataFilter = string.Empty;
        }

        #endregion Constructors

        public void SetLogger(ITraceLogger logger)
        {
            if (logger != null)
            {
                this.logger = logger;
            }
        }

        public void SetReportVisitor(ReportVisitor visitor)
        {
            if (visitor != null)
            {
                this.visitor = visitor;
            }
        }
    }
}