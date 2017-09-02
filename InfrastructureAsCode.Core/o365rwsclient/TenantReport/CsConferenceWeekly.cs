using System;
using System.Xml;
using System.Xml.Serialization;

namespace InfrastructureAsCode.Core.o365rwsclient.TenantReport
{
    [Serializable]
    public class CsConferenceWeekly : TenantReportObject
    {
        [XmlElement]
        public System.Int64 ApplicationSharingConferences { get; set; }

        [XmlElement]
        public System.Int64 AVConferences { get; set; }
        [XmlElement]
        public System.Int64 ID { get; set; }
        [XmlElement]
        public System.Int64 IMConferences { get; set; }
        [XmlElement]
        public System.Int64 TelephonyConferences { get; set; }
        [XmlElement]
        public System.Int64 TotalConferences { get; set; }
        [XmlElement]
        public System.Int64 WebConferences { get; set; }

        public override void LoadFromXml(XmlNode node)
        {
            base.LoadFromXml(node);

            ApplicationSharingConferences = Utils.TryParseInt64(base.TryGetValue("ApplicationSharingConferences"), 0);
            AVConferences = Utils.TryParseInt64(base.TryGetValue("AVConferences"), 0);
            ID = Utils.TryParseInt64(base.TryGetValue("ID"), 0);
            IMConferences = Utils.TryParseInt64(base.TryGetValue("IMConferences"), 0);
            TelephonyConferences = Utils.TryParseInt64(base.TryGetValue("TelephonyConferences"), 0);
            TotalConferences = Utils.TryParseInt64(base.TryGetValue("TotalConferences"), 0);
            WebConferences = Utils.TryParseInt64(base.TryGetValue("WebConferences"), 0);
        }
    }
}