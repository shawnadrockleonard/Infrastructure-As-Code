using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Extensions
{
    public static class StringExtensions
    {
        /// <summary>
        /// returns string value as DateTime
        /// </summary>
        /// <param name="columnValue"></param>
        /// <returns></returns>
        public static DateTime ToDate(this string columnValue)
        {
            DateTime fieldValue = DateTime.Now.AddYears(-2);
            if (!string.IsNullOrEmpty(columnValue)
                && DateTime.TryParse(columnValue, out fieldValue))
            {
            }
            return fieldValue;
        }

        /// <summary>
        /// returns string value as DateTime
        /// </summary>
        /// <param name="columnValue"></param>
        /// <returns></returns>
        public static Nullable<DateTime> ToDateTime(this string columnValue)
        {
            DateTime fieldValue = DateTime.Now.AddYears(-2);
            if (!string.IsNullOrEmpty(columnValue)
                && DateTime.TryParse(columnValue, out fieldValue))
            {
                return fieldValue;
            }
            return null;
        }

        public static Nullable<DateTime> ToNullableDatetime(this string dateTimeValue, Nullable<DateTime> defaultValue = null)
        {
            var model = (defaultValue.HasValue ? defaultValue : default(Nullable<DateTime>));
            try
            {
                if (!string.IsNullOrEmpty(dateTimeValue))
                {
                    model = DateTime.Parse(dateTimeValue);
                }
            }
            catch (Exception ex)
            {
                Trace.TraceWarning("Failed to grab {0} conversion {1}", dateTimeValue, ex.Message);
            }
            return model;
        }

        public static Int32 ToInt32(this string numberValue, Nullable<Int32> defaultValue = null)
        {
            var model = (defaultValue.HasValue ? defaultValue.Value : 0);
            try
            {
                if (!string.IsNullOrEmpty(numberValue))
                {
                    model = Int32.Parse(numberValue);
                }
            }
            catch (Exception ex)
            {
                Trace.TraceWarning("Failed to grab {0} conversion {1}", numberValue, ex.Message);
            }
            return model;
        }

        public static Int64 ToInt64(this double value, Nullable<Int64> defaultValue = null)
        {
            var model = (defaultValue.HasValue ? defaultValue.Value : 0);
            try
            {
                model = (Int64)value;
            }
            catch (Exception ex)
            {
                Trace.TraceWarning("Failed to grab {0} conversion {1}", value, ex.Message);
            }
            return model;
        }

        /// <summary>
        /// Take a double and return decimal
        /// </summary>
        /// <param name="value"></param>
        /// <param name="defaultValue"></param>
        /// <returns></returns>
        public static decimal TryParseDecimal(this double value, decimal defaultValue = 0)
        {
            decimal result = 0;
            try
            {
                result = (decimal)value;
                return result;
            }
            catch { }

            return defaultValue;
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="value"></param>
        /// <param name="defaultValue"></param>
        /// <returns></returns>
        public static decimal TryParseGB(this double value, decimal defaultValue = 0)
        {
            decimal result = 0;
            try
            {
                result = ((decimal)value / 1024);
                return result;
            }
            catch { }

            return defaultValue;
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="value"></param>
        /// <param name="defaultValue"></param>
        /// <returns></returns>
        public static decimal TryParseTB(this double value, decimal defaultValue = 0)
        {
            decimal result = 0;
            try
            {
                result = ((decimal)value / 1048576);
                return result;
            }
            catch { }

            return defaultValue;
        }

        /// <summary>
        /// Removes the XML encoded characters
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        public static string UnescapeXml(this string s)
        {
            string unxml = s;
            if (!string.IsNullOrEmpty(unxml))
            {
                // replace entities with literal values
                unxml = unxml.Replace("&apos;", "'");
                unxml = unxml.Replace("&quot;", "\"");
                unxml = unxml.Replace("&gt;", "&gt;");
                unxml = unxml.Replace("&lt;", "&lt;");
                unxml = unxml.Replace("&amp;", "&");
            }
            return unxml;
        }

        /// <summary>
        /// Replaces HTML characters to ensure XML format
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        public static string EscapeXml(this string s)
        {
            string unxml = s;
            if (!string.IsNullOrEmpty(unxml))
            {
                // replace entities with literal values
                unxml = unxml.Replace("'", "&apos;" );
                unxml = unxml.Replace("&", "&amp;");
            }
            return unxml;
        }
    }
}
