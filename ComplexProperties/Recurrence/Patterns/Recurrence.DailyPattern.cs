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