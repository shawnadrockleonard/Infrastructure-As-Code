using System;

namespace InfrastructureAsCode.Powershell.Models
{
    public class CalloutLinkModel
    {
        public string DocId { get; set; }

        public string DocIdUrl { get; set; }

        public string EmbeddedUrl { get; set; }

        public DateTime Modified { get; set; }

        public string EditorEmail { get; set; }

        public string FileUrl { get; set; }

        public string Title { get; set; }

        public int Id { get; set; }

    }
}