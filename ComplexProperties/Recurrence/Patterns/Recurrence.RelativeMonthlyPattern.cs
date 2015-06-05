/*
 * Exchange Web Services Managed API
 *
 * Copyright (c) Microsoft Corporation
 * All rights reserved.
 *
 * MIT License
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this
 * software and associated documentation files (the "Software"), to deal in the Software
 * without restriction, including without limitation the rights to use, copy, modify, merge,
 * publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
 * to whom the Software is furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all copies or
 * substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
 * INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
 * PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
 * FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
 * OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
 * DEALINGS IN THE SOFTWARE.
 */

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <content>
    /// Contains nested type Recurrence.RelativeMonthlyPattern.
    /// </content>
    public abstract partial class Recurrence
    {
        /// <summary>
        /// Represents a recurrence pattern where each occurrence happens on a relative day a specific number of months
        /// after the previous one.
        /// </summary>
        public sealed class RelativeMonthlyPattern : IntervalPattern
        {
            private DayOfTheWeek? dayOfTheWeek;
            private DayOfTheWeekIndex? dayOfTheWeekIndex;

            /// <summary>
            /// Initializes a new instance of the <see cref="RelativeMonthlyPattern"/> class.
            /// </summary>
            public RelativeMonthlyPattern()
                : base()
            {
            }

            /// <summary>
            /// Initializes a new instance of the <see cref="RelativeMonthlyPattern"/> class.
            /// </summary>
            /// <param name="startDate">The date and time when the recurrence starts.</param>
            /// <param name="interval">The number of months between each occurrence.</param>
            /// <param name="dayOfTheWeek">The day of the week each occurrence happens.</param>
            /// <param name="dayOfTheWeekIndex">The relative position of the day within the month.</param>
            public RelativeMonthlyPattern(
                DateTime startDate,
                int interval,
                DayOfTheWeek dayOfTheWeek,
                DayOfTheWeekIndex dayOfTheWeekIndex)
                : base(startDate, interval)
            {
                this.DayOfTheWeek = dayOfTheWeek;
                this.DayOfTheWeekIndex = dayOfTheWeekIndex;
            }

            /// <summary>
            /// Gets the name of the XML element.
            /// </summary>
            /// <value>The name of the XML element.</value>
            internal override string XmlElementName
            {
                get { return XmlElementNames.RelativeMonthlyRecurrence; }
            }

            /// <summary>
            /// Write properties to XML.
            /// </summary>
            /// <param name="writer">The writer.</param>
            internal override void InternalWritePropertiesToXml(EwsServiceXmlWriter writer)
            {
                base.InternalWritePropertiesToXml(writer);

                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.DaysOfWeek,
                    this.DayOfTheWeek);

                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.DayOfWeekIndex,
                    this.DayOfTheWeekIndex);
            }

            /// <summary>
            /// Tries to read element from XML.
            /// </summary>
            /// <param name="reader">The reader.</param>
            /// <returns>True if appropriate element was read.</returns>
            internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
            {
                if (base.TryReadElementFromXml(reader))
                {
                    return true;
                }
                else
                {
                    switch (reader.LocalName)
                    {
                        case XmlElementNames.DaysOfWeek:
                            this.dayOfTheWeek = reader.ReadElementValue<DayOfTheWeek>();
                            return true;
                        case XmlElementNames.DayOfWeekIndex:
                            this.dayOfTheWeekIndex = reader.ReadElementValue<DayOfTheWeekIndex>();
                            return true;
                        default:
                            return false;
                    }
                }
            }

            /// <summary>
            /// Validates this instance.
            /// </summary>
            internal override void InternalValidate()
            {
                base.InternalValidate();

                if (!this.dayOfTheWeek.HasValue)
                {
                    throw new ServiceValidationException(Strings.DayOfTheWeekMustBeSpecifiedForRecurrencePattern);
                }

                if (!this.dayOfTheWeekIndex.HasValue)
                {
                    throw new ServiceValidationException(Strings.DayOfWeekIndexMustBeSpecifiedForRecurrencePattern);
                }
            }

            /// <summary>
            /// Checks if two recurrence objects are identical. 
            /// </summary>
            /// <param name="otherRecurrence">The recurrence to compare this one to.</param>
            /// <returns>true if the two recurrences are identical, false otherwise.</returns>
            public override bool IsSame(Recurrence otherRecurrence)
            {
                RelativeMonthlyPattern otherRelativeMonthlyPattern = (RelativeMonthlyPattern)otherRecurrence;

                return base.IsSame(otherRecurrence) &&
                       this.dayOfTheWeek == otherRelativeMonthlyPattern.dayOfTheWeek &&
                       this.dayOfTheWeekIndex == otherRelativeMonthlyPattern.dayOfTheWeekIndex;
            }

            /// <summary>
            /// Gets or sets the relative position of the day specified in DayOfTheWeek within the month.
            /// </summary>
            public DayOfTheWeekIndex DayOfTheWeekIndex
            {
                get { return this.GetFieldValueOrThrowIfNull<DayOfTheWeekIndex>(this.dayOfTheWeekIndex, "DayOfTheWeekIndex"); }
                set { this.SetFieldValue<DayOfTheWeekIndex?>(ref this.dayOfTheWeekIndex, value); }
            }

            /// <summary>
            /// The day of the week when each occurrence happens.
            /// </summary>
            public DayOfTheWeek DayOfTheWeek
            {
                get { return this.GetFieldValueOrThrowIfNull<DayOfTheWeek>(this.dayOfTheWeek, "DayOfTheWeek"); }
                set { this.SetFieldValue<DayOfTheWeek?>(ref this.dayOfTheWeek, value); }
            }
        }
    }
}