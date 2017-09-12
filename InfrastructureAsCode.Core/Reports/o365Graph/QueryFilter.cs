using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Reports.o365Graph
{
    public class QueryFilter
    {
        public ReportUsageTypeEnum O365ReportType { get; set; }

        public ReportUsageViewTypeEnum ViewType { get; set; }

        public Nullable<ReportUsagePeriodEnum> O365Period { get; set; }

        public Nullable<DateTime> Date { get; set; }


        /// <summary>
        /// #Build the request URL and invoke
        ///     Sample: OneDriveActivity(view='Detail',period='D7')/content
        /// </summary>
        /// <param name="graphUrl">Represents the Graph URL for Usage Reporting which should have two parameters {0}{1}</param>
        /// <returns></returns>
        public Uri ToUrl(string graphUrl)
        {
            var str = string.Empty;

            // We always have a view to start with that
            var parameterset = string.Format("view='{0}',", ViewType.ToString("f"));

            // If period is specified then add that to the parameters unless it is not supported
            if (!Date.HasValue && O365Period.HasValue && O365ReportType != ReportUsageTypeEnum.Office365Activations)
            {
                str = string.Format("period='{0}',", O365Period.Value.ToString("f"));
                parameterset += str;
            }

            // If the date is specified then add that to the parameters unless it is not supported
            if (Date.HasValue 
                && !(O365ReportType == ReportUsageTypeEnum.MailboxUsage 
                  || O365ReportType == ReportUsageTypeEnum.Office365Activations 
                  || O365ReportType == ReportUsageTypeEnum.SfbOrganizerActivity))
            {
                str = string.Format("date='{0}'", Date.Value.ToString("yyyy-MM-dd"));
                parameterset += str;
            }

            // #Trim a trailing comma off the ParameterSet if needed
            parameterset = parameterset.TrimEnd(new char[] { ',' });

            var uri = new Uri(string.Format(graphUrl, O365ReportType, parameterset));
            return uri;
        }
    }
}
