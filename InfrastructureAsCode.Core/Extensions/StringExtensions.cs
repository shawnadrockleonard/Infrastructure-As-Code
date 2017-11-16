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
        public static DateTime ToDateTime(this string columnValue)
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
        /// <param name="dateTimeValue"></param>
        /// <returns></returns>
        public static Nullable<DateTime> ToNullableDatetime(this string dateTimeValue, Nullable<DateTime> defaultValue = null)
        {
            var returnDate = (defaultValue.HasValue ? defaultValue : default(Nullable<DateTime>));

            if (!string.IsNullOrEmpty(dateTimeValue))
            {
                if (DateTime.TryParse(dateTimeValue, out DateTime returnDateTimeValue))
                {
                    returnDate = returnDateTimeValue;
                }
            }
            return returnDate;
        }

        /// <summary>
        /// Parses the date from the string value or uses the default value if parsing fails
        /// </summary>
        /// <param name="value"></param>
        /// <param name="defaultValue"></param>
        /// <returns></returns>
        public static DateTime TryParseDateTime(this string value, DateTime defaultValue)
        {
            if (DateTime.TryParse(value, out DateTime result))
            {
                return result;
            }
            else
            {
                return defaultValue;
            }
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
        /// Will parse the string value into a small integer
        /// </summary>
        /// <param name="value"></param>
        /// <param name="defaultValue"></param>
        /// <returns>small integer</returns>
        /// <remarks>A default value will be returned if the parse fails</remarks>
        public static int TryParseInt(string value, int defaultValue)
        {
            if (int.TryParse(value, out int result))
            {
                return result;
            }
            else
            {
                return defaultValue;
            }
        }

        /// <summary>
        /// Will parse the string value into a large integer
        /// </summary>
        /// <param name="value"></param>
        /// <param name="defaultValue"></param>
        /// <returns>large integer</returns>
        /// <remarks>A default value will be returned if the parse fails</remarks>
        public static Int64 TryParseInt64(string value, Int64 defaultValue)
        {
            if (Int64.TryParse(value, out long result))
            {
                return result;
            }
            else
            {
                return defaultValue;
            }
        }

        /// <summary>
        /// Will parse the string value into a floating point number
        /// </summary>
        /// <param name="value"></param>
        /// <param name="defaultValue"></param>
        /// <returns></returns>
        public static double TryParseDouble(string value, double defaultValue)
        {
            if (double.TryParse(value, out double result))
            {
                return result;
            }
            else
            {
                return defaultValue;
            }
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
        /// Converts bytes into MegaBytes
        /// </summary>
        /// <param name="totalBytes"></param>
        /// <param name="defaultValue"></param>
        /// <returns></returns>
        public static long TryParseMB(this long totalBytes, long defaultValue = 0)
        {
            long result = 0;
            try
            {
                result = (long)(totalBytes / Math.Pow(1024, 2));
                return result;
            }
            catch { }

            return defaultValue;
        }

        /// <summary>
        /// Converts Bytes into GigaBytes
        /// </summary>
        /// <param name="totalBytes"></param>
        /// <param name="defaultValue"></param>
        /// <returns></returns>
        public static long TryParseGB(this long totalBytes, long defaultValue = 0)
        {
            long result = 0;
            try
            {
                result = (long)(totalBytes / (Math.Pow(1024, 3)));
                return result;
            }
            catch { }

            return defaultValue;
        }

        /// <summary>
        /// Converts Bytes into TeraBytes
        /// </summary>
        /// <param name="totalBytes"></param>
        /// <param name="defaultValue"></param>
        /// <returns></returns>
        public static long TryParseTB(this long totalBytes, long defaultValue = 0)
        {
            long result = 0;
            try
            {
                result = (long)(totalBytes / Math.Pow(1024, 4));
                return result;
            }
            catch { }

            return defaultValue;
        }

        /// <summary>
        /// Parses the string guid into a valid Guid or uses the Default Value
        /// </summary>
        /// <param name="value"></param>
        /// <param name="defaultValue"></param>
        /// <returns></returns>
        public static Guid TryParseGuid(this string value, Guid defaultValue)
        {
            Guid result;
            if (Guid.TryParse(value, out result))
            {
                return result;
            }

            return defaultValue;
        }

        /// <summary>
        /// Will parse the string value into a boolean
        /// </summary>
        /// <param name="value"></param>
        /// <param name="defaultValue"></param>
        /// <returns></returns>
        public static bool TryParseBoolean(this string value, bool defaultValue)
        {
            if (bool.TryParse(value, out bool result))
            {
                return result;
            }
            else
            {
                return defaultValue;
            }
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
        /// <param name="escapeQuotes">(OPTIONAL) if supplied it will XML encode double quotes</param>
        /// <returns></returns>
        public static string EscapeXml(this string s, bool escapeQuotes = false)
        {
            string unxml = s;
            if (!string.IsNullOrEmpty(unxml))
            {
                // replace entities with literal values
                if (escapeQuotes)
                {
                    unxml = unxml.Replace("\"", "&quot;").Replace(@"""", "&quot;");
                }
                unxml = unxml.Replace("'", "&apos;");
                unxml = unxml.Replace("&", "&amp;");
            }
            return unxml;
        }
    }
}
