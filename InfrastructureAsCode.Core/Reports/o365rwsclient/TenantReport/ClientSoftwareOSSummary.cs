﻿using System;
using System.Xml;
using System.Xml.Serialization;
using Utils = InfrastructureAsCode.Core.Extensions.StringExtensions;

namespace InfrastructureAsCode.Core.Reports.o365rwsclient.TenantReport
{
    [Serializable]
    public class ClientSoftwareOSSummary : TenantReportObject
    {
        [XmlElement]
        public string Name { get; set; }

        [XmlElement]
        public string Version { get; set; }

        [XmlElement]
        public System.Int64 Count { get; set; }

        [XmlElement]
        public string Category { get; set; }

        [XmlElement]
        public int DisplayOrder { get; set; }

        public override void LoadFromXml(XmlNode node)
        {
            base.LoadFromXml(node);

            Name = base.TryGetValue("Name");
            Version = base.TryGetValue("Version");
            Count = Utils.TryParseInt64(base.TryGetValue("Count"), 0);
            Category = base.TryGetValue("Category");
            DisplayOrder = Utils.TryParseInt(base.TryGetValue("DisplayOrder"), 0);
        }
    }
}