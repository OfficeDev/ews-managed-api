// ---------------------------------------------------------------------------
// <copyright file="Recurrence.YearlyPattern.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the Recurrence.YearlyPattern class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <content>
    /// Contains nested type Recurrence.YearlyPattern.
    /// </content>
    public abstract partial class Recurrence
    {
        /// <summary>
        /// Represents a recurrence pattern where each occurrence happens on a specific day every year.
        /// </summary>
        public sealed class YearlyPattern : Recurrence
        {
            private Month? month;
            private int? dayOfMonth;

            /// <summary>
            /// Initializes a new instance of the <see cref="YearlyPattern"/> class.
            /// </summary>
            public YearlyPattern()
                : base()
            {
            }

            /// <summary>
            /// Initializes a new instance of the <see cref="YearlyPattern"/> class.
            /// </summary>
            /// <param name="startDate">The date and time when the recurrence starts.</param>
            /// <param name="month">The month of the year each occurrence happens.</param>
            /// <param name="dayOfMonth">The day of the month each occurrence happens.</param>
            public YearlyPattern(
                DateTime startDate,
                Month month,
                int dayOfMonth)
                : base(startDate)
            {
                this.Month = month;
                this.DayOfMonth = dayOfMonth;
            }

            /// <summary>
            /// Gets the name of the XML element.
            /// </summary>
            /// <value>The name of the XML element.</value>
            internal override string XmlElementName
            {
                get { return XmlElementNames.AbsoluteYearlyRecurrence; }
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
                    XmlElementNames.DayOfMonth,
                    this.DayOfMonth);

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
                jsonPattern.Add(XmlElementNames.DayOfMonth, this.DayOfMonth);
                jsonPattern.Add(XmlElementNames.Month, this.Month);

                return jsonPattern;
            }

            /// <summary>
            /// Tries to read element from XML.
            /// </summary>
            /// <param name="reader">The reader.</param>
            /// <returns>True if element was read</returns>
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
                        case XmlElementNames.DayOfMonth:
                            this.dayOfMonth = reader.ReadElementValue<int>();
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
                        case XmlElementNames.DayOfMonth:
                            this.dayOfMonth = jsonProperty.ReadAsInt(key);
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
            /// Validates this instance.
            /// </summary>
            internal override void InternalValidate()
            {
                base.InternalValidate();

                if (!this.month.HasValue)
                {
                    throw new ServiceValidationException(Strings.MonthMustBeSpecifiedForRecurrencePattern);
                }

                if (!this.dayOfMonth.HasValue)
                {
                    throw new ServiceValidationException(Strings.DayOfMonthMustBeSpecifiedForRecurrencePattern);
                }
            }

            /// <summary>
            /// Gets or sets the month of the year when each occurrence happens.
            /// </summary>
            public Month Month
            {
                get { return this.GetFieldValueOrThrowIfNull<Month>(this.month, "Month"); }
                set { this.SetFieldValue<Month?>(ref this.month, value); }
            }

            /// <summary>
            /// Gets or sets the day of the month when each occurrence happens. DayOfMonth must be between 1 and 31.
            /// </summary>
            public int DayOfMonth
            {
                get
                {
                    return this.GetFieldValueOrThrowIfNull<int>(this.dayOfMonth, "DayOfMonth");
                }

                set
                {
                    if (value < 1 || value > 31)
                    {
                        throw new ArgumentOutOfRangeException("DayOfMonth", Strings.DayOfMonthMustBeBetween1And31);
                    }

                    this.SetFieldValue<int?>(ref this.dayOfMonth, value);
                }
            }
        }
    }
}