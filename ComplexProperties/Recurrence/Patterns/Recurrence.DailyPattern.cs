// ---------------------------------------------------------------------------
// <copyright file="Recurrence.DailyPattern.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the Recurrence.DailyPattern class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <content>
    /// Contains nested type Recurrence.DailyPattern.
    /// </content>
    public abstract partial class Recurrence
    {
        /// <summary>
        /// Represents a recurrence pattern where each occurrence happens a specific number of days after the previous one.
        /// </summary>
        public sealed class DailyPattern : IntervalPattern
        {
            /// <summary>
            /// Gets the name of the XML element.
            /// </summary>
            /// <value>The name of the XML element.</value>
            internal override string XmlElementName
            {
                get { return XmlElementNames.DailyRecurrence; }
            }

            /// <summary>
            /// Initializes a new instance of the <see cref="DailyPattern"/> class.
            /// </summary>
            public DailyPattern()
                : base()
            {
            }

            /// <summary>
            /// Initializes a new instance of the <see cref="DailyPattern"/> class.
            /// </summary>
            /// <param name="startDate">The date and time when the recurrence starts.</param>
            /// <param name="interval">The number of days between each occurrence.</param>
            public DailyPattern(DateTime startDate, int interval)
                : base(startDate, interval)
            {
            }
        }
    }
}