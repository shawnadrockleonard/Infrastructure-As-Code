using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Models
{
    /// <summary>
    /// Represents a Recycle Bin Item
    /// </summary>
    public class SPRecycleBinElement
    {
        public string Title { get; set; }

        public Guid Id { get; set; }

        public long FileSize { get; set; }

        public string LeafName { get; set; }

        public string DirName { get; set; }

        public string Author { get; set; }
        public string AuthorEmail { get; set; }

        public string DeletedBy { get; set; }
        public string DeletedByEmail { get; set; }

        public DateTime Deleted { get; set; }

        /// <summary>
        /// What type of item is this
        /// </summary>
        public RecycleBinItemType FileType { get; set; }


        public RecycleBinItemState FileState { get; set; }
    }
}
