// ---------------------------------------------------------------------------
// <copyright file="Recurrence.RelativeMonthlyPattern.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the Recurrence.RelativeMonthlyPattern class.</summary>
//-----------------------------------------------------------------------

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
            /// Patterns to json.
            /// </summary>
            /// <param name="service">The service.</param>
            /// <returns></returns>
            internal override JsonObject PatternToJson(ExchangeService service)
            {
                JsonObject jsonPattern = base.PatternToJson(service);

                jsonPattern.Add(XmlElementNames.DaysOfWeek, this.DayOfTheWeek);
                jsonPattern.Add(XmlElementNames.DayOfWeekIndex, this.DayOfTheWeekIndex);

                return jsonPattern;
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
                        default:
                            break;
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