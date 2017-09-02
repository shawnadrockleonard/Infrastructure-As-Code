using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Utilities
{
    public enum PartitionKeyTypeEnum
    {
        /// <summary>
        /// each logger gets his own partition in Table Storage
        /// </summary>
        LoggerName,

        /// <summary>
        /// order by Date Reverse to see the latest items first
        /// <see cref="http://gauravmantri.com/2012/02/17/effective-way-of-fetching-diagnostics-data-from-windows-azure-diagnostics-table-hint-use-partitionkey/"/>
        /// </summary>
        DateReverse
    }
}
