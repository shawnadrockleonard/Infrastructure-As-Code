using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Extensions
{
    /// <summary>
    /// Provides methods to error check and extract field values
    /// </summary>
    public static class ListItemExtensions
    {

        /// <summary>
        /// Grabs column value and if populated returns string value
        /// </summary>
        /// <param name="requestItem"></param>
        /// <param name="columnName"></param>
        /// <returns></returns>
        public static string RetrieveListItemValue(this ListItem requestItem, string columnName)
        {
            var fieldItemValue = requestItem[columnName];
            if (fieldItemValue != null)
            {
                return fieldItemValue.ToString();
            }
            return string.Empty;
        }

        /// <summary>
        /// Grabs the column value and if populated returns the FieldUserValue object otherwise null
        /// </summary>
        /// <param name="requestItem"></param>
        /// <param name="columnName"></param>
        /// <returns></returns>
        public static string[] RetrieveListItemChoiceValues(this ListItem requestItem, string columnName)
        {
            var fieldItemValue = requestItem[columnName];
            if (fieldItemValue != null)
            {
                return (string[])fieldItemValue;
            }
            return new string[0];
        }

        /// <summary>
        /// Grabs the column value and if populated returns the FieldUserValue object otherwise null
        /// </summary>
        /// <param name="requestItem"></param>
        /// <param name="columnName"></param>
        /// <returns></returns>
        public static FieldMultiChoice RetrieveListItemAsChoice(this ListItem requestItem, string columnName)
        {
            var fieldItemValue = requestItem[columnName];
            if (fieldItemValue != null)
            {
                return (FieldMultiChoice)fieldItemValue;
            }
            return null;
        }

        /// <summary>
        /// Grabs the column value and if populated returns the FieldUserValue object otherwise null
        /// </summary>
        /// <param name="requestItem"></param>
        /// <param name="columnName"></param>
        /// <returns></returns>
        public static FieldUserValue RetrieveListItemUserValue(this ListItem requestItem, string columnName)
        {
            var fieldItemValue = requestItem[columnName];
            if (fieldItemValue != null)
            {
                return (FieldUserValue)fieldItemValue;
            }
            return null;
        }

        public static string ToUserValue(this FieldUserValue fieldItemValue)
        {
            if (fieldItemValue != null)
            {
                return fieldItemValue.LookupValue;
            }
            return string.Empty;
        }

        public static string ToUserEmailValue(this FieldUserValue fieldItemValue)
        {
            if (fieldItemValue != null)
            {
                return fieldItemValue.Email;
            }
            return string.Empty;
        }

        public static string ToUserEmailValue(this Microsoft.SharePoint.Client.User fieldItemValue)
        {
            if (fieldItemValue != null)
            {
                return fieldItemValue.Email;
            }
            return string.Empty;
        }

        /// <summary>
        /// Grabs the column value and if populated returns the FieldUserValue object otherwise null
        /// </summary>
        /// <param name="requestItem"></param>
        /// <param name="columnName"></param>
        /// <returns></returns>
        public static FieldUserValue[] RetrieveListItemUserValues(this ListItem requestItem, string columnName)
        {
            var fieldItemValue = requestItem[columnName];
            if (fieldItemValue != null)
            {
                return fieldItemValue as FieldUserValue[];
            }
            return null;
        }

        /// <summary>
        /// Parse the field user values into an array of strings
        /// </summary>
        /// <param name="fieldItemValue"></param>
        /// <returns></returns>
        public static IEnumerable<string> ToUserValues(this FieldUserValue[] fieldItemValue)
        {
            if (fieldItemValue != null)
            {
                return fieldItemValue.Select(s => s.ToUserValue());
            }
            return new string[0];
        }

        /// <summary>
        /// Grabs column value and if populated returns field lookup value object
        /// </summary>
        /// <param name="requestItem"></param>
        /// <param name="columnName"></param>
        /// <returns></returns>
        public static FieldLookupValue RetrieveListItemValueAsLookup(this ListItem requestItem, string columnName)
        {
            var fieldItemValue = requestItem[columnName];
            if (fieldItemValue != null)
            {
                return (FieldLookupValue)fieldItemValue;
            }
            return null;
        }

        /// <summary>
        /// Grabs column value and if populated returns field lookup value object
        /// </summary>
        /// <param name="requestItem"></param>
        /// <param name="columnName"></param>
        /// <returns></returns>
        public static FieldLookupValue[] RetrieveListItemValueAsLookups(this ListItem requestItem, string columnName)
        {
            var fieldItemValue = requestItem[columnName];
            if (fieldItemValue != null)
            {
                return (FieldLookupValue[])fieldItemValue;
            }
            return null;
        }

        public static string ToLookupValue(this FieldLookupValue fieldItemValue)
        {
            if (fieldItemValue != null)
            {
                return fieldItemValue.LookupValue;
            }
            return string.Empty;
        }


        /// <summary>
        /// Grabs column value and if populated returns field hyperlink value object
        /// </summary>
        /// <param name="requestItem"></param>
        /// <param name="columnName"></param>
        /// <returns></returns>
        public static FieldUrlValue RetrieveListItemValueAsHyperlink(this ListItem requestItem, string columnName)
        {
            var fieldItemValue = requestItem[columnName];
            if (fieldItemValue != null)
            {
                return (FieldUrlValue)fieldItemValue;
            }
            return null;
        }
    }
}
