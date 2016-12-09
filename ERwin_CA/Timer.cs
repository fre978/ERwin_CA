using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace ERwin_CA
{
    /// <summary>
    /// Manage time measurement.
    /// </summary>
    static class Timer
    {
        private static DateTime StartTime;
        private static DateTime EndTime;

        public static void SetFirstTime(this DateTime value)
        {
            StartTime = value;
        }

        public static DateTime GetFirstTime()
        {
            return StartTime;
        }

        public static void SetSecondTime(this DateTime value)
        {
            EndTime = value;
        }

        public static DateTime GetSecondTime()
        {
            return EndTime;
        }

        public static TimeSpan GetTimeLapse(DateTime first, DateTime second)
        {
            TimeSpan res;
            if (DateTime.Compare(first, second) == -1)
            {
                res = second - first;
                return res;
            }
            else
            {
                res = first - second;
                return res;
            }
        }

        public static string GetTimeLapseFormatted(DateTime first, DateTime second)
        {
            TimeSpan ts = GetTimeLapse(first, second);
            string result = string.Format("{0:D2}h:{1:D2}m:{2:D2}s:{3:D3}ms",
                                    ts.Hours,
                                    ts.Minutes,
                                    ts.Seconds,
                                    ts.Milliseconds);
            return result;
        }

        public static String GetTimestampPrecision(this DateTime value)
        {
            return value.ToString("HH:mm:ss:fff");
        }
        public static String GetTimestampMinute(this DateTime value)
        {
            return value.ToString("yyyyMMdd_HH:mm");
        }

        public static String GetTimestampDay(this DateTime value)
        {
            return value.ToString("yyyyMMdd");
        }

        public static String GetTimestampFolder(this DateTime value)
        {
            return value.ToString("yyyyMMdd_HHmm");
        }

    }
}
