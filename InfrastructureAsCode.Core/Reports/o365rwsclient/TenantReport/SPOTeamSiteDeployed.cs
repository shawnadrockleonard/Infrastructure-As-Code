using System.Xml;
using System.Xml.Serialization;
using Utils = InfrastructureAsCode.Core.Extensions.StringExtensions;

namespace InfrastructureAsCode.Core.Reports.o365rwsclient.TenantReport
{
    public class SPOTeamSiteDeployed : TenantReportObject
    {
        [XmlElement]
        public System.Int64 ID
        {
            get;
            set;
        }

        [XmlElement]
        public double Active
        {
            get;
            set;
        }

        [XmlElement]
        public double Inactive
        {
            get;
            set;
        }

        public override void LoadFromXml(XmlNode node)
        {
            base.LoadFromXml(node);
            ID = Utils.TryParseInt64(base.TryGetValue("ID"), 0);
            Active = Utils.TryParseDouble(base.TryGetValue("Active"), 0);
            Inactive = Utils.TryParseDouble(base.TryGetValue("Inactive"), 0);
        }
    }
}