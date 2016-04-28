using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace Microsoft.IaC.Core.Utilities
{
    /// <summary>
    /// PRovides an easy way to read an xml config file
    /// </summary>
    public class ConfigurationReader
    {
        private readonly XElement[] appSettings;

        private readonly XElement[] connectionSettings;

        public ConfigurationReader(string path)
        {
            var document = XElement.Load(path);
            appSettings = document.Descendants("appSettings").Descendants("add").ToArray();
            connectionSettings = document.Descendants("connectionStrings").Descendants("add").ToArray();
        }


        public string GetAppSetting(string index)
        {
            return appSettings.FirstOrDefault(e => e.Attribute("key").Value == index).Attribute("value").Value;
        }

        public string GetConnectionSetting(string index)
        {
            return appSettings.FirstOrDefault(e => e.Attribute("name").Value == index).Attribute("connectionString").Value;
        }


    }
}
