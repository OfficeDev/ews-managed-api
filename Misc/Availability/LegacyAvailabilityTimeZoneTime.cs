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
// <summary>Defines the LegacyAvailabilityTimeZoneTime class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a custom time zone time change. 
    /// </summary>
    internal sealed class LegacyAvailabilityTimeZoneTime : ComplexProperty
    {
        private TimeSpan delta;
        private int year;
        private int month;
        private int dayOrder;
        private DayOfTheWeek dayOfTheWeek;
        private TimeSpan timeOfDay;

        /// <summary>
        /// Initializes a new instance of the <see cref="LegacyAvailabilityTimeZoneTime"/> class.
        /// </summary>
        internal LegacyAvailabilityTimeZoneTime()
            : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="LegacyAvailabilityTimeZoneTime"/> class.
        /// </summary>
        /// <param name="transitionTime">The transition time used to initialize this instance.</param>
        /// <param name="delta">The offset used to initialize this instance.</param>
        internal LegacyAvailabilityTimeZoneTime(TimeZoneInfo.TransitionTime transitionTime, TimeSpan delta)
            : this()
        {
            this.delta = delta;

            if (transitionTime.IsFixedDateRule)
            {
                // TimeZoneInfo doesn't support an actual year. Fixed date transitions occur at the same
                // date every year the adjustment rule the transition belongs to applies. The best thing
                // we can do here is use the current year.
                this.year = DateTime.Today.Year;
                this.month = transitionTime.Month;
                this.dayOrder = transitionTime.Day;
                this.timeOfDay = transitionTime.TimeOfDay.TimeOfDay;
            }
            else
            {
                // For floating rules, the mapping is direct.
                this.year = 0;
                this.month = transitionTime.Month;
                this.dayOfTheWeek = EwsUtilities.SystemToEwsDayOfTheWeek(transitionTime.DayOfWeek);
                this.dayOrder = transitionTime.Week;
                this.timeOfDay = transitionTime.TimeOfDay.TimeOfDay;
            }
        }

        /// <summary>
        /// Converts this instance to TimeZoneInfo.TransitionTime.
        /// </summary>
        /// <returns>A TimeZoneInfo.TransitionTime</returns>
        internal TimeZoneInfo.TransitionTime ToTransitionTime()
        {
            if (this.year == 0)
            {
                return TimeZoneInfo.TransitionTime.CreateFloatingDateRule(
                    new DateTime(
                        DateTime.MinValue.Year,
                        DateTime.MinValue.Month,
                        DateTime.MinValue.Day,
                        this.timeOfDay.Hours,
                        this.timeOfDay.Minutes,
                        this.timeOfDay.Seconds),
                    this.month,
                    this.dayOrder,
                    EwsUtilities.EwsToSystemDayOfWeek(this.dayOfTheWeek));
            }
            else
            {
                return TimeZoneInfo.TransitionTime.CreateFixedDateRule(
                    new DateTime(this.timeOfDay.Ticks),
                    this.month,
                    this.dayOrder);
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
                case XmlElementNames.Bias:
                    this.delta = TimeSpan.FromMinutes(reader.ReadElementValue<int>());
                    return true;
                case XmlElementNames.Time:
                    this.timeOfDay = TimeSpan.Parse(reader.ReadElementValue());
                    return true;
                case XmlElementNames.DayOrder:
                    this.dayOrder = reader.ReadElementValue<int>();
                    return true;
                case XmlElementNames.DayOfWeek:
                    this.dayOfTheWeek = reader.ReadElementValue<DayOfTheWeek>();
                    return true;
                case XmlElementNames.Month:
                    this.month = reader.ReadElementValue<int>();
                    return true;
                case XmlElementNames.Year:
                    this.year = reader.ReadElementValue<int>();
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        /// Loads from json.
        /// </summary>
        /// <param name="jsonProperty">The json property.</param>
        /// <param name="service"></param>
        internal override void LoadFromJson(JsonObject jsonProperty, ExchangeService service)
        {
            foreach (string key in jsonProperty.Keys)
            {
                switch (key)
                {
                    case XmlElementNames.Bias:
                        this.delta = TimeSpan.FromMinutes(jsonProperty.ReadAsInt(key));
                        break;
                    case XmlElementNames.Time:
                        this.timeOfDay = TimeSpan.Parse(jsonProperty.ReadAsString(key));
                        break;
                    case XmlElementNames.DayOrder:
                        this.dayOrder = jsonProperty.ReadAsInt(key);
                        break;
                    case XmlElementNames.DayOfWeek:
                        this.dayOfTheWeek = jsonProperty.ReadEnumValue<DayOfTheWeek>(key);
                        break;
                    case XmlElementNames.Month:
                        this.month = jsonProperty.ReadAsInt(key);
                        break;
                    case XmlElementNames.Year:
                        this.year = jsonProperty.ReadAsInt(key);
                        break;
                    default:
                        break;
                }
            }
        }

        /// <summary>
        /// Writes the elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteElementValue(
                XmlNamespace.Types,
                XmlElementNames.Bias,
                (int)this.delta.TotalMinutes);

            writer.WriteElementValue(
                XmlNamespace.Types,
                XmlElementNames.Time,
                EwsUtilities.TimeSpanToXSTime(this.timeOfDay));

            writer.WriteElementValue(
                XmlNamespace.Types,
                XmlElementNames.DayOrder,
                this.dayOrder);

            writer.WriteElementValue(
                XmlNamespace.Types,
                XmlElementNames.Month,
                (int)this.month);

            // Only write DayOfWeek if this is a recurring time change
            if (this.Year == 0)
            {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.DayOfWeek,
                    this.dayOfTheWeek);
            }

            // Only emit year if it's non zero, otherwise AS returns "Request is invalid"
            if (this.Year != 0)
            {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.Year,
                    this.Year);
            }
        }

        /// <summary>
        /// Serializes the property to a Json value.
        /// </summary>
        /// <param name="service"></param>
        /// <returns>
        /// A Json value (either a JsonObject, an array of Json values, or a Json primitive)
        /// </returns>
        internal override object InternalToJson(ExchangeService service)
        {
            JsonObject jsonProperty = new JsonObject();

            jsonProperty.Add(
                XmlElementNames.Bias,
                (int)this.delta.TotalMinutes);

            jsonProperty.Add(
                XmlElementNames.Time,
                EwsUtilities.TimeSpanToXSTime(this.timeOfDay));

            jsonProperty.Add(
                XmlElementNames.DayOrder,
                this.dayOrder);

            jsonProperty.Add(
                XmlElementNames.Month,
                (int)this.month);

            // Only write DayOfWeek if this is a recurring time change
            if (this.Year == 0)
            {
                jsonProperty.Add(
                    XmlElementNames.DayOfWeek,
                    this.dayOfTheWeek);
            }

            // Only emit year if it's non zero, otherwise AS returns "Request is invalid"
            if (this.Year != 0)
            {
                jsonProperty.Add(
                    XmlElementNames.Year,
                    this.Year);
            }

            return jsonProperty;
        }

        /// <summary>
        /// Gets if current time presents DST transition time
        /// </summary>
        internal bool HasTransitionTime
        {
            get { return this.month >= 1 && this.month <= 12; }
        }

        /// <summary>
        /// Gets or sets the delta.
        /// </summary>
        internal TimeSpan Delta
        {
            get { return this.delta; }
            set { this.delta = value; }
        }

        /// <summary>
        /// Gets or sets the time of day.
        /// </summary>
        internal TimeSpan TimeOfDay
        {
            get { return this.timeOfDay; }
            set { this.timeOfDay = value; }
        }

        /// <summary>
        /// Gets or sets a value that represents:
        /// - The day of the month when Year is non zero,
        /// - The index of the week in the month if Year is equal to zero.
        /// </summary>
        internal int DayOrder
        {
            get { return this.dayOrder; }
            set { this.dayOrder = value; }
        }

        /// <summary>
        /// Gets or sets the month.
        /// </summary>
        internal int Month
        {
            get { return this.month; }
            set { this.month = value; }
        }

        /// <summary>
        /// Gets or sets the day of the week.
        /// </summary>
        internal DayOfTheWeek DayOfTheWeek
        {
            get { return this.dayOfTheWeek; }
            set { this.dayOfTheWeek = value; }
        }

        /// <summary>
        /// Gets or sets the year. If Year is 0, the time change occurs every year according to a recurring pattern;
        /// otherwise, the time change occurs at the date specified by Day, Month, Year.
        /// </summary>
        internal int Year
        {
            get { return this.year; }
            set { this.year = value; }
        }
    }
}