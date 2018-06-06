using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Reports
{
    /// <summary>
    /// Provides a mechanism to parse various data values into a specific type
    /// </summary>
    public static class ReportDataExtensions
    {
        /// <summary>
        /// Parse Large Int the value into a the specific power (i.e. convert Bytes into Mb, Gb, Tb, Pb)
        /// </summary>
        /// <param name="webValue"></param>
        /// <param name="power"></param>
        /// <param name="defaultValue"></param>
        /// <returns></returns>
        public static decimal ParseInt64DecimalByPower(this Int64 webValue, int power, decimal defaultValue = 0)
        {
            decimal result = defaultValue;
            try
            {
                result = (decimal)(webValue / (Math.Pow(1024, power)));
            }
            catch { }
            return result;
        }

        /// <summary>
        /// Parse Large Int the value into a the specific power (i.e. convert Bytes into Mb, Gb, Tb, Pb)
        /// </summary>
        /// <param name="webValue"></param>
        /// <param name="power"></param>
        /// <param name="defaultValue"></param>
        /// <returns></returns>
        public static decimal ParseLongDecimalByPower(this long webValue, int power, decimal defaultValue = 0)
        {
            decimal result = defaultValue;
            try
            {
                result = (decimal)(webValue / (Math.Pow(1024, power)));
            }
            catch { }
            return result;
        }

        /// <summary>
        /// Parse Double the value into a the specific power (i.e. convert Bytes into Mb, Gb, Tb, Pb)
        /// </summary>
        /// <param name="webValue"></param>
        /// <param name="power"></param>
        /// <param name="defaultValue"></param>
        /// <returns></returns>
        public static decimal ParseDoubleDecimalByPower(this double webValue, int power, decimal defaultValue = 0)
        {
            decimal result = defaultValue;
            try
            {
                result = (decimal)(webValue / (Math.Pow(1024, power)));
            }
            catch { }
            return result;
        }

        /// <summary>
        /// Parse Large Int the value into a the specific power (i.e. convert Bytes into Mb, Gb, Tb, Pb)
        /// </summary>
        /// <param name="webValue"></param>
        /// <param name="power"></param>
        /// <param name="defaultValue"></param>
        /// <returns></returns>
        public static decimal ParseInt64DecimalByPower(this Nullable<Int64> webValue, int power, decimal defaultValue = 0)
        {
            decimal result = defaultValue;
            try
            {
                result = (decimal)(webValue / (Math.Pow(1024, power)));
            }
            catch { }
            return result;
        }

        /// <summary>
        /// Parse Large Int the value into a the specific power (i.e. convert Bytes into Mb, Gb, Tb, Pb)
        /// </summary>
        /// <param name="webValue"></param>
        /// <param name="power"></param>
        /// <param name="defaultValue"></param>
        /// <returns></returns>
        public static decimal ParseLongDecimalByPower(this Nullable<long> webValue, int power, decimal defaultValue = 0)
        {
            decimal result = defaultValue;
            try
            {
                result = (decimal)(webValue / (Math.Pow(1024, power)));
            }
            catch { }
            return result;
        }

        /// <summary>
        /// Parse Double Floating Point the value into a the specific power (i.e. convert Bytes into Mb, Gb, Tb, Pb)
        /// </summary>
        /// <param name="webValue"></param>
        /// <param name="power"></param>
        /// <param name="defaultValue"></param>
        /// <returns></returns>
        public static decimal ParseDoubleDecimalByPower(this Nullable<double> webValue, int power, decimal defaultValue = 0)
        {
            decimal result = defaultValue;
            try
            {
                result = (decimal)(webValue / (Math.Pow(1024, power)));
            }
            catch { }
            return result;
        }

        /// <summary>
        /// Converts nullable into default value
        /// </summary>
        /// <param name="webValue"></param>
        /// <param name="defaultValue"></param>
        /// <returns></returns>
        public static Int64 ParseInt64Default(this Nullable<Int64> webValue, Int64 defaultValue = 0)
        {
            Int64 result = defaultValue;
            try
            {
                result = (Int64)webValue;
            }
            catch { }

            return result;
        }

        /// <summary>
        /// Converts nullable into default value
        /// </summary>
        /// <param name="webValue"></param>
        /// <param name="defaultValue"></param>
        /// <returns></returns>
        public static long ParseDefault(this Nullable<long> webValue, long defaultValue = 0)
        {
            long result = defaultValue;
            try
            {
                result = (long)webValue;
            }
            catch { }

            return result;
        }

    }
}
