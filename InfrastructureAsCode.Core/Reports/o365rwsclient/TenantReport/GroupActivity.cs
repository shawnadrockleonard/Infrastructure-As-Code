using System.Xml;
using System.Xml.Serialization;
using Utils = InfrastructureAsCode.Core.Extensions.StringExtensions;

namespace InfrastructureAsCode.Core.Reports.o365rwsclient.TenantReport
{
    public class GroupActivity : TenantReportObject
    {
        [XmlElement]
        public int GroupCreated
        {
            get;
            set;
        }

        [XmlElement]
        public int GroupDeleted
        {
            get;
            set;
        }

        public override void LoadFromXml(XmlNode node)
        {
            base.LoadFromXml(node);
            GroupCreated = Utils.TryParseInt(base.TryGetValue("GroupCreated"), 0);
            GroupDeleted = Utils.TryParseInt(base.TryGetValue("GroupDeleted"), 0);
        }
    }
}