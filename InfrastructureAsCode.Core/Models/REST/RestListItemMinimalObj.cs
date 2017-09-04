using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Models.REST
{
    public class RestListItemMinimalObj : RestBaseMinimalObject, IRestListItemObj
    {
        public int Id { get; set; }

        public string ContentTypeId { get; set; }

        public string Title { get; set; }

        public Microsoft.SharePoint.Client.FileSystemObjectType FileSystemObjectType { get; set; }

        public DateTime Modified { get; set; }

        public DateTime Created { get; set; }

        public int AuthorId { get; set; }

        public int EditorId { get; set; }

        public string OData__UIVersionString { get; set; }

        public bool Attachments { get; set; }

        public Guid GUID { get; set; }
    }
}
