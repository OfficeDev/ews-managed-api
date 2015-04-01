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
    /// Contains nested type Recurrence.RelativeYearlyPattern.
    /// </content>
    public abstract partial class Recurrence
    {
        /// <summary>
        /// Represents a recurrence pattern where each occurrence happens on a relative day every year.
        /// </summary>
        public sealed class RelativeYearlyPattern : Recurrence
        {
            private DayOfTheWeek? dayOfTheWeek;
            private DayOfTheWeekIndex? dayOfTheWeekIndex;
            private Month? month;

            /// <summary>
            /// Gets the name of the XML element.
            /// </summary>
            /// <value>The name of the XML element.</value>
            internal override string XmlElementName
            {
                get { return XmlElementNames.RelativeYearlyRecurrence; }
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

                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.Month,
                    this.Month);
            }

            /// <summary>
            /// Patterns to json.
            /// </summary>
            /// <param name="service">The service.</param>
            /// <returns></returns>
            internal override JsonObject PatternToJson(ExchangeService service)
            {
                JsonObject jsonPattern = new JsonObject();

                jsonPattern.AddTypeParameter(this.XmlElementName);

                jsonPattern.Add(XmlElementNames.DaysOfWeek, this.DayOfTheWeek);
                jsonPattern.Add(XmlElementNames.DayOfWeekIndex, this.DayOfTheWeekIndex);
                jsonPattern.Add(XmlElementNames.Month, this.Month);

                return jsonPattern;
            }

            /// <summary>
            /// Tries to read element from XML.
            /// </summary>
            /// <param name="reader">The reader.</param>
            /// <returns>True if element was read.</returns>
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
                        case XmlElementNames.Month:
                            this.month = reader.ReadElementValue<Month>();
                            return true;
                        default:
                            return false;
                    }
                }
            }

            /// <summary>
            /// Loads from json.
            /// </summary>
            /// <param name="jsonProperty">The json property.</param>
            /// <param name="service">The service.</param>
            internal override void LoadFromJson(JsonObject jsonProperty, ExchangeService service)
            {
                base.LoadFromJson(jsonProperty, service);

                foreach (string key in jsonProperty.Keys)
                {
                    switch (key)
                    {
                        case XmlElementNames.DaysOfWeek:
                            this.dayOfTheWeek = jsonProperty.ReadEnumValue<DayOfTheWeek>(key);
                            break;
                        case XmlElementNames.DayOfWeekIndex:
                            this.dayOfTheWeekIndex = jsonProperty.ReadEnumValue<DayOfTheWeekIndex>(key);
                            break;
                        case XmlElementNames.Month:
                            this.month = jsonProperty.ReadEnumValue<Month>(key);
                            break;
                        default:
                            break;
                    }
                }
            }

            /// <summary>
            /// Initializes a new instance of the <see cref="RelativeYearlyPattern"/> class.
            /// </summary>
            public RelativeYearlyPattern()
                : base()
            {
            }

            /// <summary>
            /// Initializes a new instance of the <see cref="RelativeYearlyPattern"/> class.
            /// </summary>
            /// <param name="startDate">The date and time when the recurrence starts.</param>
            /// <param name="month">The month of the year each occurrence happens.</param>
            /// <param name="dayOfTheWeek">The day of the week each occurrence happens.</param>
            /// <param name="dayOfTheWeekIndex">The relative position of the day within the month.</param>
            public RelativeYearlyPattern(
                DateTime startDate,
                Month month,
                DayOfTheWeek dayOfTheWeek,
                DayOfTheWeekIndex dayOfTheWeekIndex)
                : base(startDate)
            {
                this.Month = month;
                this.DayOfTheWeek = dayOfTheWeek;
                this.DayOfTheWeekIndex = dayOfTheWeekIndex;
            }

            /// <summary>
            /// Validates this instance.
            /// </summary>
            internal override void InternalValidate()
            {
                base.InternalValidate();

                if (!this.dayOfTheWeekIndex.HasValue)
                {
                    throw new ServiceValidationException(Strings.DayOfWeekIndexMustBeSpecifiedForRecurrencePattern);
                }

                if (!this.dayOfTheWeek.HasValue)
                {
                    throw new ServiceValidationException(Strings.DayOfTheWeekMustBeSpecifiedForRecurrencePattern);
                }

                if (!this.month.HasValue)
                {
                    throw new ServiceValidationException(Strings.MonthMustBeSpecifiedForRecurrencePattern);
                }
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
            /// Gets or sets the day of the week when each occurrence happens.
            /// </summary>
            public DayOfTheWeek DayOfTheWeek
            {
                get { return this.GetFieldValueOrThrowIfNull<DayOfTheWeek>(this.dayOfTheWeek, "DayOfTheWeek"); }
                set { this.SetFieldValue<DayOfTheWeek?>(ref this.dayOfTheWeek, value); }
            }

            /// <summary>
            /// Gets or sets the month of the year when each occurrence happens.
            /// </summary>
            public Month Month
            {
                get { return this.GetFieldValueOrThrowIfNull<Month>(this.month, "Month"); }
                set { this.SetFieldValue<Month?>(ref this.month, value); }
            }
        }
    }
}