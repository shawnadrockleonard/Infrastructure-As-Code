using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Powershell.Extensions
{
    /// <summary>
    /// PS Object expression extensions
    /// </summary>
    public static class ObjectExtensions
    {

        public static string GetPSObjectValue(this PSMemberInfoCollection<PSPropertyInfo> infoProperties, string propertyName)
        {
            var resultValue = string.Empty;
            if (infoProperties[propertyName] != null && infoProperties[propertyName].Value != null)
            {
                resultValue = infoProperties[propertyName].Value.ToString();
            }
            return resultValue;
        }


        public static string GetPropertyValue(this List<KeyValuePair<string, string>> profileProperties, string propertyName, Type valueType = null)
        {
            var property = profileProperties.FirstOrDefault(f => f.Key.Equals(propertyName, StringComparison.CurrentCultureIgnoreCase));
            if (!property.Equals(default(KeyValuePair<string, string>)))
            {
                if (valueType != null && valueType == typeof(System.Guid))
                {
                    var propValue = new Guid(property.Value);
                    return propValue.ToString("D");
                }
                return property.Value;
            }
            return string.Empty;
        }

    }
}
