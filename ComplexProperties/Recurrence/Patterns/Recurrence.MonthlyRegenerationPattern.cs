// ---------------------------------------------------------------------------
// <copyright file="Recurrence.MonthlyRegenerationPattern.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the Recurrence.MonthlyRegenerationPattern class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <content>
    /// Contains nested type Recurrence.MonthlyRegenerationPattern.
    /// </content>
    public abstract partial class Recurrence
    {
        /// <summary>
        /// Represents a regeneration pattern, as used with recurring tasks, where each occurrence happens
        /// a specified number of months after the previous one is completed.
        /// </summary>
        public sealed class MonthlyRegenerationPattern : IntervalPattern
        {
            /// <summary>
            /// Initializes a new instance of the <see cref="MonthlyRegenerationPattern"/> class.
            /// </summary>
            public MonthlyRegenerationPattern()
                : base()
            {
            }

            /// <summary>
            /// Initializes a new instance of the <see cref="MonthlyRegenerationPattern"/> class.
            /// </summary>
            /// <param name="startDate">The date and time when the recurrence starts.</param>
            /// <param name="interval">The number of months between previous and next occurrences.</param>
            public MonthlyRegenerationPattern(DateTime startDate, int interval)
                : base(startDate, interval)
            {
            }

            /// <summary>
            /// Gets the name of the XML element.
            /// </summary>
            /// <value>The name of the XML element.</value>
            internal override string XmlElementName
            {
                get { return XmlElementNames.MonthlyRegeneration; }
            }

            /// <summary>
            /// Gets a value indicating whether this instance is regeneration pattern.
            /// </summary>
            /// <value>
            ///     <c>true</c> if this instance is regeneration pattern; otherwise, <c>false</c>.
            /// </value>
            internal override bool IsRegenerationPattern
            {
                get { return true; }
            }
        }
    }
}