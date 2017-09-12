using System.Xml;
using System.Xml.Serialization;
using Utils = InfrastructureAsCode.Core.Extensions.StringExtensions;

namespace InfrastructureAsCode.Core.Reports.o365rwsclient.TenantReport
{
    public class SPOActiveUser : TenantReportObject
    {
        [XmlElement]
        public System.Int64 ID
        {
            get;
            set;
        }

        [XmlElement]
        public double UniqueUsers
        {
            get;
            set;
        }

        [XmlElement]
        public double LicensesAssigned
        {
            get;
            set;
        }

        [XmlElement]
        public double LicensesAcquired
        {
            get;
            set;
        }

        public double TotalUsers
        {
            get;
            set;
        }

        public override void LoadFromXml(XmlNode node)
        {
            base.LoadFromXml(node);
            ID = Utils.TryParseInt64(base.TryGetValue("ID"), 0);
            UniqueUsers = Utils.TryParseDouble(base.TryGetValue("UniqueUsers"), 0);
            LicensesAssigned = Utils.TryParseDouble(base.TryGetValue("LicensesAssigned"), 0);
            LicensesAcquired = Utils.TryParseDouble(base.TryGetValue("LicensesAcquired"), 0);
            TotalUsers = Utils.TryParseDouble(base.TryGetValue("TotalUsers"), 0);
        }
    }
}