﻿using System.Xml;
using System.Xml.Serialization;
using Utils = InfrastructureAsCode.Core.Extensions.StringExtensions;

namespace InfrastructureAsCode.Core.Reports.o365rwsclient.TenantReport
{
    public class ConnectionByClientType : TenantReportObject
    {
        [XmlElement]
        public string ClientType { get; set; }

        [XmlElement]
        public System.Int64 Count { get; set; }

        public override void LoadFromXml(XmlNode node)
        {
            base.LoadFromXml(node);

            ClientType = base.TryGetValue("ClientType");
            Count = Utils.TryParseInt64(base.TryGetValue("Count"), 0);
        }
    }
}