using log4net.Core;
using Microsoft.Azure;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Utilities
{
    internal static class LoggingEventExtensions
    {
        internal const string AzureConnectionStringNotSpecified = "Azure ConnectionString not specified";
        internal const string TableNameNotSpecified = "Table name not specified";
        internal const string ElasticTableEntity_GetEntityProperty_not_supported__0__for__1_ = "not supported {0} for {1}";

        internal static string MakeRowKey(this LoggingEvent loggingEvent)
        {
            return string.Format("{0:D19}.{1}",
                 DateTime.MaxValue.Ticks - loggingEvent.TimeStamp.Ticks, Guid.NewGuid().ToString().ToLower());
        }

        internal static string MakePartitionKey(this LoggingEvent loggingEvent, PartitionKeyTypeEnum partitionKeyType)
        {
            switch (partitionKeyType)
            {
                case PartitionKeyTypeEnum.LoggerName:
                    return loggingEvent.LoggerName;
                case PartitionKeyTypeEnum.DateReverse:
                    // substract from DateMaxValue the Tick Count of the current hour
                    // so a Table Storage Partition spans an hour
                    return string.Format("{0:D19}",
                        (DateTime.MaxValue.Ticks -
                         loggingEvent.TimeStamp.Date.AddHours(loggingEvent.TimeStamp.Hour).Ticks + 1));
                default:
                    // ReSharper disable once NotResolvedInText
                    throw new ArgumentOutOfRangeException("PartitionKeyType", partitionKeyType, null);
            }
        }

        internal static IEnumerable<IEnumerable<TSource>> Batch<TSource>(this IEnumerable<TSource> source, int size)
        {
            TSource[] bucket = null;
            var count = 0;

            foreach (var item in source)
            {
                if (bucket == null)
                    bucket = new TSource[size];

                bucket[count++] = item;
                if (count != size)
                    continue;

                yield return bucket;

                bucket = null;
                count = 0;
            }

            if (bucket != null && count > 0)
                yield return bucket.Take(count);
        }

        /// <summary>
        /// Attempt to retrieve the connection string using ConfigurationManager 
        /// with CloudConfigurationManager as fallback
        /// </summary>
        /// <param name="connectionStringName">The name of the connection string to retrieve</param>
        /// <returns></returns>
        internal static string GetConnectionString(this string connectionStringName)
        {
            // Attempt to retrieve the connection string using the regular ConfigurationManager first
            var config = ConfigurationManager.ConnectionStrings[connectionStringName];
            if (config != null)
            {
                return config.ConnectionString;
            }

            // Fallback to CloudConfigurationManager in case we're running as a worker/web role
            var azConfig = CloudConfigurationManager.GetSetting(connectionStringName);
            if (!string.IsNullOrWhiteSpace(azConfig))
            {
                return azConfig;
            }

            // Connection string not found, throw exception to notify the user
            throw new ApplicationException(AzureConnectionStringNotSpecified);
        }
    }
}
