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
// <summary>Defines the TimeChangeRecurrence class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Text;

    /// <summary>
    /// Represents a recurrence pattern for a time change in a time zone.
    /// </summary>
    internal sealed class TimeChangeRecurrence : ComplexProperty
    {
        private DayOfTheWeek? dayOfTheWeek;
        private DayOfTheWeekIndex? dayOfTheWeekIndex;
        private Month? month;

        /// <summary>
        /// Initializes a new instance of the <see cref="TimeChangeRecurrence"/> class.
        /// </summary>
        public TimeChangeRecurrence()
            : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="TimeChangeRecurrence"/> class.
        /// </summary>
        /// <param name="dayOfTheWeekIndex">The index of the day in the month at which the time change occurs.</param>
        /// <param name="dayOfTheWeek">The day of the week the time change occurs.</param>
        /// <param name="month">The month the time change occurs.</param>
        public TimeChangeRecurrence(
            DayOfTheWeekIndex dayOfTheWeekIndex,
            DayOfTheWeek dayOfTheWeek,
            Month month)
            : this()
        {
            this.dayOfTheWeekIndex = dayOfTheWeekIndex;
            this.dayOfTheWeek = dayOfTheWeek;
            this.month = month;
        }

        /// <summary>
        /// Gets or sets the index of the day in the month at which the time change occurs.
        /// </summary>
        public DayOfTheWeekIndex? DayOfTheWeekIndex
        {
            get { return this.dayOfTheWeekIndex; }
            set { this.SetFieldValue<DayOfTheWeekIndex?>(ref this.dayOfTheWeekIndex, value); }
        }

        /// <summary>
        /// Gets or sets the day of the week the time change occurs.
        /// </summary>
        public DayOfTheWeek? DayOfTheWeek
        {
            get { return this.dayOfTheWeek; }
            set { this.SetFieldValue<DayOfTheWeek?>(ref this.dayOfTheWeek, value); }
        }

        /// <summary>
        /// Gets or sets the month the time change occurs.
        /// </summary>
        public Month? Month
        {
            get { return this.month; }
            set { this.SetFieldValue<Month?>(ref this.month, value); }
        }

        /// <summary>
        /// Writes elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            if (this.DayOfTheWeek.HasValue)
            {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.DaysOfWeek,
                    this.DayOfTheWeek.Value);
            }

            if (this.dayOfTheWeekIndex.HasValue)
            {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.DayOfWeekIndex,
                    this.DayOfTheWeekIndex.Value);
            }

            if (this.Month.HasValue)
            {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.Month,
                    this.Month.Value);
            }
        }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            switch (reader.LocalName)
            {
                case XmlElementNames.DaysOfWeek:
                    this.dayOfTheWeek = reader.ReadElementValue<DayOfTheWeek>();
                    return true;
                case XmlElementNames.DayOfWeekIndex:
                    this.dayOfTheWeekIndex = reader.ReadElementValue<DayOfTheWeekIndex>();
                    return true;
                case XmlElementNames.Month:
                    this.month = reader.ReadElementValue<Month>();
                    return true;
                default:
                    return false;
            }
        }
    }
}