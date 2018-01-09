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
        public static decimal ParseDecimalByPower(this long webValue, int power, decimal defaultValue = 0)
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
        public static decimal ParseDecimalByPower(this double webValue, int power, decimal defaultValue = 0)
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
        public static decimal ParseDecimalByPower(this Nullable<long> webValue, int power, decimal defaultValue = 0)
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
        public static decimal ParseDecimalByPower(this Nullable<double> webValue, int power, decimal defaultValue = 0)
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
