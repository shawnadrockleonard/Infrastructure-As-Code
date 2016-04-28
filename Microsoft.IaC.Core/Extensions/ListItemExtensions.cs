using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.IaC.Core.Extensions
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
        public static FieldUserValue RetrieveListItemUserValue(this ListItem requestItem, string columnName)
        {
            var fieldItemValue = requestItem[columnName];
            if (fieldItemValue != null)
            {
                return (FieldUserValue)fieldItemValue;
            }
            return null;
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
    }
}
