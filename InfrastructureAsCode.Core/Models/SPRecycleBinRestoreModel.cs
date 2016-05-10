using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Models
{
    public class SPRecycleBinRestoreModel : SPRecycleBinElement
    {
        public SPRecycleBinRestoreModel()
        {

        }

        /// <summary>
        /// Extend the bin model with the paging set in which the item resides
        /// </summary>
        public string PagingInfo { get; set; }

    }
}
