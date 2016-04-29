using OfficeDevPnP.Core.UPAWebService;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IaC.Core.Extensions
{
    public static class UserExtensions
    {

        public static string RetrieveUserProperty(this GetUserProfileByIndexResult UserProfileResult, string propertyName)
        {
            var propertyValue = string.Empty;
            try
            {
                var Prop = UserProfileResult.UserProfile.FirstOrDefault(w => w.Name == propertyName);
                if (Prop != null && Prop.Values.Length > 0)
                {
                    propertyValue = Prop.Values[0].Value.ToString();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.TraceWarning("Property name {0} not found with message {1}", propertyName, ex.Message);
            }
            return propertyValue;
        }

        /// <summary>
        /// Enumerates the collection to find the property name and return the first value
        /// </summary>
        /// <param name="UserProfileResult"></param>
        /// <param name="propertyName"></param>
        /// <returns></returns>
        public static string RetrieveUserProperty(this PropertyData[] UserProfileResult, string propertyName)
        {
            var propertyValue = string.Empty;
            try
            {
                var Prop = UserProfileResult.FirstOrDefault(w => w.Name == propertyName);
                if (Prop != null && Prop.Values.Length > 0)
                {
                    propertyValue = Prop.Values[0].Value.ToString();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.TraceWarning("Property name {0} not found with message {1}", propertyName, ex.Message);
            }
            return propertyValue;
        }
    }
}
