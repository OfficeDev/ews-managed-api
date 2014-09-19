// ---------------------------------------------------------------------------
// <copyright file="Recurrence.DailyRegenerationPattern.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the Recurrence.DailyRegenerationPattern class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <content>
    /// Contains nested type Recurrence.DailyRegenerationPattern.
    /// </content>
    public abstract partial class Recurrence
    {
        /// <summary>
        /// Represents a regeneration pattern, as used with recurring tasks, where each occurrence happens a specified number of days after the previous one is completed.
        /// </summary>
        public sealed class DailyRegenerationPattern : IntervalPattern
        {
            /// <summary>
            /// Initializes a new instance of the <see cref="DailyRegenerationPattern"/> class.
            /// </summary>
            public DailyRegenerationPattern()
                : base()
            {
            }

            /// <summary>
            /// Initializes a new instance of the <see cref="DailyRegenerationPattern"/> class.
            /// </summary>
            /// <param name="startDate">The date and time when the recurrence starts.</param>
            /// <param name="interval">The number of days between the current occurrence and the next, after the current occurrence is completed.</param>
            public DailyRegenerationPattern(DateTime startDate, int interval)
                : base(startDate, interval)
            {
            }

            /// <summary>
            /// Gets the name of the XML element.
            /// </summary>
            /// <value>The name of the XML element.</value>
            internal override string XmlElementName
            {
                get { return XmlElementNames.DailyRegeneration; }
            }

            /// <summary>
            /// Gets a value indicating whether this instance is a regeneration pattern.
            /// </summary>
            /// <value><c>true</c> if this instance is a regeneration pattern; otherwise, <c>false</c>.</value>
            internal override bool IsRegenerationPattern
            {
                get { return true; }
            }
        }
    }
}