// ---------------------------------------------------------------------------
// <copyright file="DayOfTheWeek.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the DayOfTheWeek enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    
    /// <summary>
    /// Specifies the day of the week.
    /// </summary>
    /// <remarks>
    /// For the standard days of the week (Sunday, Monday...) the DayOfTheWeek enum value is the same as the System.DayOfWeek 
    /// enum type. These values can be safely cast between the two enum types. The special days of the week (Day, Weekday and
    /// WeekendDay) are used for monthly and yearly recurrences and cannot be cast to System.DayOfWeek values.
    /// </remarks>
    public enum DayOfTheWeek
    {
        /// <summary>
        /// Sunday
        /// </summary>
        Sunday = DayOfWeek.Sunday,

        /// <summary>
        /// Monday
        /// </summary>
        Monday = DayOfWeek.Monday,

        /// <summary>
        /// Tuesday
        /// </summary>
        Tuesday = DayOfWeek.Tuesday,

        /// <summary>
        /// Wednesday
        /// </summary>
        Wednesday = DayOfWeek.Wednesday,

        /// <summary>
        /// Thursday
        /// </summary>
        Thursday = DayOfWeek.Thursday,

        /// <summary>
        /// Friday
        /// </summary>
        Friday = DayOfWeek.Friday,

        /// <summary>
        /// Saturday
        /// </summary>
        Saturday = DayOfWeek.Saturday,

        /// <summary>
        /// Any day of the week
        /// </summary>
        Day,

        /// <summary>
        /// Any day of the usual business week (Monday-Friday)
        /// </summary>
        Weekday,

        /// <summary>
        /// Any weekend day (Saturday or Sunday)
        /// </summary>
        WeekendDay,
    }
}
