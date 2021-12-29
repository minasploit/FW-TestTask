using System;
using System.Collections.Generic;

namespace FW_TestTask
{
    public static class Extensions
    {
        public static string GetFriendlyName(this TimeSpan time)
        {
            if (time.TotalHours >= 1)
            {
                return $"{(int)time.TotalHours}:{time.Minutes:00} hours";
            }
            if (time.TotalMinutes >= 1)
            {
                return $"{time.Minutes}:{time.Seconds:00} minutes";
            }
            return time.TotalSeconds >= 1 ? $"{time.Seconds} seconds" : $"{time.Milliseconds} milliseconds";
        }
        
        public static IEnumerable<DateTime> DatesBetween(this DateTime from, DateTime to)
        {
            for(var day = from.Date; day.Date <= to.Date; day = day.AddDays(1))
                yield return day;
        }
    }
}