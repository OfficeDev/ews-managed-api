using System;
using System.Reflection;

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Utility class for declaring time zone related extension methods.
    /// </summary>
    public static class TimeZoneInfoExtensionMethods
    {
        /// <summary>
        /// Extension method to return the internal BaseUtcOffsetDelta property value from an AdjustmentRule.
        /// </summary>
        /// <param name="adjustmentRule">The adjustement rule whose BaseUtcOffsetDelta value should be returned.</param>
        /// <returns>A TimeSpan value that reprensents the AdjustmentRule's BaseUtcOffsetDelta value.</returns>
        public static TimeSpan GetBaseUtcOffsetDelta(this TimeZoneInfo.AdjustmentRule adjustmentRule)
        {
            BindingFlags bindFlags = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;
            PropertyInfo property = typeof(TimeZoneInfo.AdjustmentRule).GetProperty("BaseUtcOffsetDelta", bindFlags);
            return (TimeSpan)property.GetValue(adjustmentRule, null);
        }

        /// <summary>
        /// Extension method to determine if two TransitionTime values reference the same day.
        /// </summary>
        /// <param name="thisTransitionTime">The first TransitionTime to compare.</param>
        /// <param name="otherTransitionTime">The second TransitionTime to compare.</param>
        /// <returns>True if both TransitionTime objects reference the same day, otherwise false.</returns>
        public static bool HasSameDate(this TimeZoneInfo.TransitionTime thisTransitionTime, TimeZoneInfo.TransitionTime otherTransitionTime)
        {
            if (thisTransitionTime.IsFixedDateRule && otherTransitionTime.IsFixedDateRule)
            {
                return (thisTransitionTime.Month == otherTransitionTime.Month)
                    && (thisTransitionTime.Day == otherTransitionTime.Day);
            }
            else if (!thisTransitionTime.IsFixedDateRule && !otherTransitionTime.IsFixedDateRule)
            {
                return (thisTransitionTime.Month == otherTransitionTime.Month)
                    && (thisTransitionTime.Week == otherTransitionTime.Week)
                    && (thisTransitionTime.DayOfWeek == otherTransitionTime.DayOfWeek);
            }
            else
            {
                return false;
            }
        }
    }
}
