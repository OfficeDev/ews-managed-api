#region License

// Exchange Web Services Managed API
// 
// Copyright (c) Microsoft Corporation
// All rights reserved. 
//
// MIT License
//
// Permission is hereby granted, free of charge, to any person obtaining a copy of this
// software and associated documentation files (the "Software"), to deal in the Software
// without restriction, including without limitation the rights to use, copy, modify, merge,
// publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
// to whom the Software is furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all copies or
// substantial portions of the Software.

// THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
// INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
// PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
// FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
// OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
// DEALINGS IN THE SOFTWARE.

#endregion

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