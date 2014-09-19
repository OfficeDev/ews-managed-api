// ---------------------------------------------------------------------------
// <copyright file="DayOfTheWeekIndex.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the DayOfTheWeekIndex enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Defines the index of a week day within a month.
    /// </summary>
    public enum DayOfTheWeekIndex
    {
        /// <summary>
        /// The first specific day of the week in the month. For example, the first Tuesday of the month. 
        /// </summary>
        First,

        /// <summary>
        /// The second specific day of the week in the month. For example, the second Tuesday of the month.
        /// </summary>
        Second,

        /// <summary>
        /// The third specific day of the week in the month. For example, the third Tuesday of the month.
        /// </summary>
        Third,

        /// <summary>
        /// The fourth specific day of the week in the month. For example, the fourth Tuesday of the month.
        /// </summary>
        Fourth,

        /// <summary>
        /// The last specific day of the week in the month. For example, the last Tuesday of the month.
        /// </summary>
        Last
    }
}
