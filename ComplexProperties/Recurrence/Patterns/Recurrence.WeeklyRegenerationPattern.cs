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
// <summary>Defines the Recurrence.WeeklyRegenerationPattern class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <content>
    /// Contains nested type Recurrence.WeeklyRegenerationPattern.
    /// </content>
    public abstract partial class Recurrence
    {
        /// <summary>
        /// Represents a regeneration pattern, as used with recurring tasks, where each occurrence happens a specified number of weeks after the previous one is completed.
        /// </summary>
        public sealed class WeeklyRegenerationPattern : IntervalPattern
        {
            /// <summary>
            /// Initializes a new instance of the <see cref="WeeklyRegenerationPattern"/> class.
            /// </summary>
            public WeeklyRegenerationPattern()
                : base()
            {
            }

            /// <summary>
            /// Initializes a new instance of the <see cref="WeeklyRegenerationPattern"/> class.
            /// </summary>
            /// <param name="startDate">The date and time when the recurrence starts.</param>
            /// <param name="interval">The number of weeks between the current occurrence and the next, after the current occurrence is completed.</param>
            public WeeklyRegenerationPattern(DateTime startDate, int interval)
                : base(startDate, interval)
            {
            }

            /// <summary>
            /// Gets the name of the XML element.
            /// </summary>
            /// <value>The name of the XML element.</value>
            internal override string XmlElementName
            {
                get { return XmlElementNames.WeeklyRegeneration; }
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