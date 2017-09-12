using System.Xml;
using System.Xml.Serialization;
using Utils = InfrastructureAsCode.Core.Extensions.StringExtensions;

namespace InfrastructureAsCode.Core.Reports.o365rwsclient.TenantReport
{
    public class MailboxActivity : TenantReportObject
    {
        [XmlElement]
        public int TotalNumberOfActiveMailboxes
        {
            get;
            set;
        }

        [XmlElement]
        public int AccountCreated
        {
            get;
            set;
        }

        [XmlElement]
        public int AccountDeleted
        {
            get;
            set;
        }

        public override void LoadFromXml(XmlNode node)
        {
            base.LoadFromXml(node);
            TotalNumberOfActiveMailboxes = Utils.TryParseInt(base.TryGetValue("TotalNumberOfActiveMailboxes"), 0);
            AccountCreated = Utils.TryParseInt(base.TryGetValue("AccountCreated"), 0);
            AccountDeleted = Utils.TryParseInt(base.TryGetValue("AccountDeleted"), 0);
        }
    }
}