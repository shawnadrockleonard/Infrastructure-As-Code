﻿using System.Xml;
using System.Xml.Serialization;
using Utils = InfrastructureAsCode.Core.Extensions.StringExtensions;

namespace InfrastructureAsCode.Core.Reports.o365rwsclient.TenantReport
{
    public class SPOSkyDriveProStorage : TenantReportObject
    {
        [XmlElement]
        public System.Int64 ID
        {
            get;
            set;
        }

        [XmlElement]
        public double Used
        {
            get;
            set;
        }

        [XmlElement]
        public double Allocated
        {
            get;
            set;
        }

        public override void LoadFromXml(XmlNode node)
        {
            base.LoadFromXml(node);
            ID = Utils.TryParseInt64(base.TryGetValue("ID"), 0);
            Used = Utils.TryParseDouble(base.TryGetValue("Used"), 0);
            Allocated = Utils.TryParseDouble(base.TryGetValue("Allocated"), 0);
        }
    }
}