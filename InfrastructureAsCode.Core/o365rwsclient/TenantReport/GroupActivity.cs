using System.Xml;
using System.Xml.Serialization;

namespace InfrastructureAsCode.Core.o365rwsclient.TenantReport
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