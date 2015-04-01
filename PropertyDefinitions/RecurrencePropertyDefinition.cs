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
    using System.Xml;

    /// <summary>
    /// Represenrs recurrence property definition.
    /// </summary>
    internal sealed class RecurrencePropertyDefinition : PropertyDefinition
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="RecurrencePropertyDefinition"/> class.
        /// </summary>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <param name="uri">The URI.</param>
        /// <param name="flags">The flags.</param>
        /// <param name="version">The version.</param>
        internal RecurrencePropertyDefinition(
            string xmlElementName,
            string uri,
            PropertyDefinitionFlags flags,
            ExchangeVersion version)
            : base(
                xmlElementName,
                uri,
                flags,
                version)
        {
        }

        /// <summary>
        /// Loads from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <param name="propertyBag">The property bag.</param>
        internal override void LoadPropertyValueFromXml(EwsServiceXmlReader reader, PropertyBag propertyBag)
        {
            reader.EnsureCurrentNodeIsStartElement(XmlNamespace.Types, XmlElementNames.Recurrence);

            Recurrence recurrence = null;

            reader.Read(XmlNodeType.Element); // This is the pattern element

            recurrence = GetRecurrenceFromString(reader.LocalName);

            recurrence.LoadFromXml(reader, reader.LocalName);

            reader.Read(XmlNodeType.Element); // This is the range element

            RecurrenceRange range = GetRecurrenceRange(reader.LocalName);

            range.LoadFromXml(reader, reader.LocalName);
            range.SetupRecurrence(recurrence);

            reader.ReadEndElementIfNecessary(XmlNamespace.Types, XmlElementNames.Recurrence);

            propertyBag[this] = recurrence;
        }

        /// <summary>
        /// Loads the property value from json.
        /// </summary>
        /// <param name="value">The JSON value.  Can be a JsonObject, string, number, bool, array, or null.</param>
        /// <param name="service">The service.</param>
        /// <param name="propertyBag">The property bag.</param>
        internal override void LoadPropertyValueFromJson(object value, ExchangeService service, PropertyBag propertyBag)
        {
            JsonObject jsonRecurrence = value as JsonObject;

            JsonObject jsonPattern = jsonRecurrence.ReadAsJsonObject(JsonNames.RecurrencePattern);            
            Recurrence recurrence = GetRecurrenceFromString(jsonPattern.ReadTypeString());
            recurrence.LoadFromJson(jsonPattern, service);

            JsonObject jsonRange = jsonRecurrence.ReadAsJsonObject(JsonNames.RecurrenceRange);
            RecurrenceRange range = GetRecurrenceRange(jsonRange.ReadTypeString());
            range.LoadFromJson(jsonRange, service);

            range.SetupRecurrence(recurrence);

            propertyBag[this] = recurrence;
        }

        /// <summary>
        /// Gets the recurrence range.
        /// </summary>
        /// <param name="recurrenceRangeString">The recurrence range string.</param>
        /// <returns></returns>
        private static RecurrenceRange GetRecurrenceRange(string recurrenceRangeString)
        {
            RecurrenceRange range;

            switch (recurrenceRangeString)
            {
                case XmlElementNames.NoEndRecurrence:
                    range = new NoEndRecurrenceRange();
                    break;
                case XmlElementNames.EndDateRecurrence:
                    range = new EndDateRecurrenceRange();
                    break;
                case XmlElementNames.NumberedRecurrence:
                    range = new NumberedRecurrenceRange();
                    break;
                default:
                    throw new ServiceXmlDeserializationException(string.Format(Strings.InvalidRecurrenceRange, recurrenceRangeString));
            }
            return range;
        }

        /// <summary>
        /// Gets the recurrence from string.
        /// </summary>
        /// <param name="recurranceString">The recurrance string.</param>
        /// <returns></returns>
        private static Recurrence GetRecurrenceFromString(string recurranceString)
        {
            Recurrence recurrence = null;

            switch (recurranceString)
            {
                case XmlElementNames.RelativeYearlyRecurrence:
                    recurrence = new Recurrence.RelativeYearlyPattern();
                    break;
                case XmlElementNames.AbsoluteYearlyRecurrence:
                    recurrence = new Recurrence.YearlyPattern();
                    break;
                case XmlElementNames.RelativeMonthlyRecurrence:
                    recurrence = new Recurrence.RelativeMonthlyPattern();
                    break;
                case XmlElementNames.AbsoluteMonthlyRecurrence:
                    recurrence = new Recurrence.MonthlyPattern();
                    break;
                case XmlElementNames.DailyRecurrence:
                    recurrence = new Recurrence.DailyPattern();
                    break;
                case XmlElementNames.DailyRegeneration:
                    recurrence = new Recurrence.DailyRegenerationPattern();
                    break;
                case XmlElementNames.WeeklyRecurrence:
                    recurrence = new Recurrence.WeeklyPattern();
                    break;
                case XmlElementNames.WeeklyRegeneration:
                    recurrence = new Recurrence.WeeklyRegenerationPattern();
                    break;
                case XmlElementNames.MonthlyRegeneration:
                    recurrence = new Recurrence.MonthlyRegenerationPattern();
                    break;
                case XmlElementNames.YearlyRegeneration:
                    recurrence = new Recurrence.YearlyRegenerationPattern();
                    break;
                default:
                    throw new ServiceXmlDeserializationException(string.Format(Strings.InvalidRecurrencePattern, recurranceString));
            }
            return recurrence;
        }

        /// <summary>
        /// Writes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="propertyBag">The property bag.</param>
        /// <param name="isUpdateOperation">Indicates whether the context is an update operation.</param>
        internal override void WritePropertyValueToXml(
            EwsServiceXmlWriter writer,
            PropertyBag propertyBag,
            bool isUpdateOperation)
        {
            Recurrence value = (Recurrence)propertyBag[this];

            if (value != null)
            {
                value.WriteToXml(writer, XmlElementNames.Recurrence);
            }
        }

        /// <summary>
        /// Writes the json value.
        /// </summary>
        /// <param name="jsonObject">The json object.</param>
        /// <param name="propertyBag">The property bag.</param>
        /// <param name="service">The service.</param>
        /// <param name="isUpdateOperation">if set to <c>true</c> [is update operation].</param>
        internal override void WriteJsonValue(JsonObject jsonObject, PropertyBag propertyBag, ExchangeService service, bool isUpdateOperation)
        {
            Recurrence value = propertyBag[this] as Recurrence;

            if (value != null)
            {
                jsonObject.Add(this.XmlElementName, value.InternalToJson(service));
            }
        }

        /// <summary>
        /// Gets the property type.
        /// </summary>
        public override Type Type
        {
            get { return typeof(Recurrence); }
        }
    }
}