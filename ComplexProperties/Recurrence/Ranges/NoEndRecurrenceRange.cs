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
// <summary>Defines the NoEndRecurrenceRange class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents recurrence range with no end date.
    /// </summary>
    internal sealed class NoEndRecurrenceRange : RecurrenceRange
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="NoEndRecurrenceRange"/> class.
        /// </summary>
        public NoEndRecurrenceRange()
            : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="NoEndRecurrenceRange"/> class.
        /// </summary>
        /// <param name="startDate">The start date.</param>
        public NoEndRecurrenceRange(DateTime startDate)
            : base(startDate)
        {
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <value>The name of the XML element.</value>
        internal override string XmlElementName
        {
            get { return XmlElementNames.NoEndRecurrence; }
        }

        /// <summary>
        /// Setups the recurrence.
        /// </summary>
        /// <param name="recurrence">The recurrence.</param>
        internal override void SetupRecurrence(Recurrence recurrence)
        {
            base.SetupRecurrence(recurrence);

            recurrence.NeverEnds();
        }
    }
}
