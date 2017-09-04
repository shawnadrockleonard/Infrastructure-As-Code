using System;
using Microsoft.SharePoint.Client;

namespace InfrastructureAsCode.Core.Models.REST
{
    public interface IRestListItemObj
    {
        bool Attachments { get; set; }

        int AuthorId { get; set; }

        string ContentTypeId { get; set; }

        DateTime Created { get; set; }

        int EditorId { get; set; }

        FileSystemObjectType FileSystemObjectType { get; set; }

        Guid GUID { get; set; }

        int Id { get; set; }

        DateTime Modified { get; set; }

        string OData__UIVersionString { get; set; }

        string Title { get; set; }
    }
}