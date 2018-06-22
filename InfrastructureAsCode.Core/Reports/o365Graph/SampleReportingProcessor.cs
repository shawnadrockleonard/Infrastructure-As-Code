using InfrastructureAsCode.Core.oAuth;
using InfrastructureAsCode.Core.Reports.o365Graph.TenantReport.Mappings;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Reports.o365Graph
{
    public class SampleReportingProcessor : IDisposable
    {
        #region Properties
 
        internal IAzureADConfig ADConfig { get; private set; }

        internal ITraceLogger Logger { get; private set; }

        public bool _disposed { get; set; }

        /// <summary>
        /// Represents the EF processing throttle
        /// </summary>
        private readonly int throttle = 50;

        #endregion


        public SampleReportingProcessor()
        {
        }

        public SampleReportingProcessor( IAzureADConfig _adconfig, ITraceLogger _logger) : this()
        {
            ADConfig = _adconfig;
            Logger = _logger;
        }

        public void ProcessOneDriveActivityFileCounts(ICollection<OneDriveActivityFileCounts> results)
        {
            // getOneDriveActivityFileCounts
            var rowprocessed = results.Count();

            var rowdx = 0; var totaldx = rowprocessed;
            foreach (var result in results)
            {
                rowdx++;
                totaldx--;

                if(rowdx % throttle == 0 || totaldx <= 0)
                {
                    Logger.LogWarning($"Now running storage processing for {rowdx} batch of records");
                }
            }

            Logger.LogInformation("Processed Activity Files {0} rows", rowprocessed);
        }

        public void ProcessOneDriveActivityUserCounts(ICollection<OneDriveActivityUserCounts> results)
        {
            var rowprocessed = results.Count();

            var rowdx = 0; var totaldx = rowprocessed;
            foreach (var result in results)
            {
                rowdx++;
                totaldx--;
            }

            Logger.LogInformation("Processed Activity Users {0} rows", rowprocessed);
        }


        public void ProcessOneDriveActivityUserDetail(ICollection<OneDriveActivityUserDetail> results)
        {
            var rowprocessed = results.Count();

            // Return the dates so we can filter the EF query to possible Key overlaps
            var maxdates = results.Select(gb => gb.LastActivityDateUTC).Distinct();

            var rowdx = 0; var totaldx = rowprocessed;
            foreach (var result in results)
            {
                rowdx++;
                totaldx--;
            }

            Logger.LogInformation("Completed processing of OneDriveActivityDetail {0} rows", rowprocessed);
        }


        public void ProcessOneDriveUsageAccountCounts(ICollection<OneDriveUsageAccountCounts> results)
        {
            var rowprocessed = results.Count();

            var rowdx = 0; var totaldx = rowprocessed;
            foreach (var result in results)
            {
                rowdx++;
                totaldx--;
            }

            Logger.LogInformation("Processed Usage Account {0} rows", rowprocessed);
        }


        public void ProcessOneDriveUsageFileCounts(ICollection<OneDriveUsageFileCounts> results)
        {
            var rowprocessed = results.Count();

            Logger.LogInformation("Processed Usage Account {0} rows", rowprocessed);
        }


        public void ProcessOneDriveUsageStorage(ICollection<OneDriveUsageStorage> results)
        {
            var rowprocessed = results.Count();

            var rowdx = 0; var totaldx = rowprocessed;
            foreach (var result in results)
            {
                rowdx++;
                totaldx--;
            }

            Logger.LogInformation("Processed Usage Storage {0} rows", rowprocessed);
        }

        public void ProcessOneDriveUsageAccountDetail(ICollection<OneDriveUsageAccountDetail> results)
        {
            var rowprocessed = results.Count();

            // Return the dates so we can filter the EF query to possible Key overlaps
            var maxdates = results.Select(gb => gb.LastActivityDateUTC).Distinct();

            var rowdx = 0; var totaldx = rowprocessed;
            foreach (var result in results)
            {
                rowdx++;
                totaldx--;
            }

            Logger.LogInformation("Completed processing of OneDriveUsageAccountDetail {0} rows", rowprocessed);
        }


        public void ProcessOffice365ActiveUserDetail(ICollection<Office365ActiveUsersUserDetail> results)
        {
            var rowprocessed = results.Count();

            var rowdx = 0; var totaldx = results.Count();
            foreach (var result in results)
            {
                rowdx++;
                totaldx--;
            }

            Logger.LogInformation("Completed processing of Office365ActiveUsersUserDetail {0} rows", rowprocessed);
        }


        public void ProcessOffice365ServicesUserCounts(ICollection<Office365ActiveUsersServicesUserCounts> results)
        {
            var rowprocessed = results.Count();

            var rowdx = 0; var totaldx = results.Count();
            foreach (var result in results)
            {
                rowdx++;
                totaldx--;
            }

            Logger.LogInformation("Completed processing of Office365ActiveUsersServicesUserCounts {0} rows", rowprocessed);
        }


        public void ProcessOffice365ActiveUserCounts(ICollection<Office365ActiveUsersUserCounts> results)
        {
            var rowprocessed = results.Count();

            var rowdx = 0; var totaldx = results.Count();
            foreach (var result in results)
            {
                rowdx++;
                totaldx--;
            }

            Logger.LogInformation("Completed processing of Office365ActiveUsersUserCounts {0} rows", rowprocessed);
        }

        public void ProcessSharePointActivityFileCounts(ICollection<SharePointActivityFileCounts> results)
        {

        }

        public void ProcessSharePointActivityPages(ICollection<SharePointActivityPages> results)
        {
        }

        public void ProcessSharePointActivityUserCounts(ICollection<SharePointActivityUserCounts> results)
        {
        }

        public void ProcessSharePointActivityUserDetail(ICollection<SharePointActivityUserDetail> results)
        {
        }

        public void ProcessSharePointSiteUsageDetail(ICollection<SharePointSiteUsageSiteDetail> results)
        {
            DateTime.TryParse("2014-01-01", out DateTime nulldefaultdate);

            // Return the dates so we can filter the EF query to possible Key overlaps
            var maxdates = results.Select(gb =>
            {
                if (gb.LastActivityDate.HasValue) return gb.LastActivityDate;
                return nulldefaultdate;
            }).Distinct();


            var rowdx = 0; var totaldx = results.Count();
            foreach (var result in results)
            {
                rowdx++;
                totaldx--;
            }

            Logger.LogInformation("Completed processing of SharePointSiteUsageDetail {0} rows", results.Count());
        }


        public void ProcessSharePointSiteUsageFileCounts(ICollection<SharePointSiteUsageFileCounts> results)
        {
        }

        public void ProcessSharePointSiteUsagePages(ICollection<SharePointSiteUsagePages> results)
        {
        }

        public void ProcessSharePointSiteUsageSiteCounts(ICollection<SharePointSiteUsageSiteCounts> results)
        {

            // Return the dates so we can filter the EF query to possible Key overlaps
            var maxdates = results.Select(gb => gb.ReportDate).Distinct();

            var rowdx = 0; var totaldx = results.Count();
            foreach (var result in results)
            {
                rowdx++;
                totaldx--;
            }

            Logger.LogInformation("Completed processing of SharePointSiteUsageSiteCounts {0} rows", results.Count());
        }


        public void ProcessSharePointSiteUsageStorage(ICollection<SharePointSiteUsageStorage> results)
        {
            // Return the dates so we can filter the EF query to possible Key overlaps
            var maxdates = results.Select(gb => gb.ReportDate).Distinct();
            
            var rowdx = 0; var totaldx = results.Count();
            foreach (var result in results)
            {
                rowdx++;
                totaldx--;
            }

            Logger.LogInformation("Completed processing of SharePointSiteUsageStorage {0} rows", results.Count());
        }



        /// <summary>
        /// Converts nullable into default value
        /// </summary>
        /// <param name="webValue"></param>
        /// <param name="defaultValue"></param>
        /// <returns></returns>
        protected long ParseDefault(Nullable<long> webValue, long defaultValue = 0)
        {
            long result = defaultValue;
            try
            {
                result = (long)webValue;
            }
            catch { }

            return result;
        }


        protected virtual void Dispose(bool disposing)
        {
            if (disposing
                && !_disposed)
            {
            }
            _disposed = true;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
    }
}
