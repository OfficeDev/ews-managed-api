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

    /// <summary>
    /// Represents the definition for the meeting time zone property.
    /// </summary>
    internal class MeetingTimeZonePropertyDefinition : PropertyDefinition
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="MeetingTimeZonePropertyDefinition"/> class.
        /// </summary>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <param name="uri">The URI.</param>
        /// <param name="flags">The flags.</param>
        /// <param name="version">The version.</param>
        internal MeetingTimeZonePropertyDefinition(
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
        internal override sealed void LoadPropertyValueFromXml(EwsServiceXmlReader reader, PropertyBag propertyBag)
        {
            MeetingTimeZone meetingTimeZone = new MeetingTimeZone();
            meetingTimeZone.LoadFromXml(reader, this.XmlElementName);

            propertyBag[AppointmentSchema.StartTimeZone] = meetingTimeZone.ToTimeZoneInfo();
        }

        /// <summary>
        /// Loads the property value from json.
        /// </summary>
        /// <param name="value">The JSON value.  Can be a JsonObject, string, number, bool, array, or null.</param>
        /// <param name="service">The service.</param>
        /// <param name="propertyBag">The property bag.</param>
        internal override void LoadPropertyValueFromJson(object value, ExchangeService service, PropertyBag propertyBag)
        {
            JsonObject jsonObject = value as JsonObject;

            if (jsonObject != null)
            {
                MeetingTimeZone meetingTimeZone = new MeetingTimeZone();
                meetingTimeZone.LoadFromJson(jsonObject, service);

                propertyBag[AppointmentSchema.StartTimeZone] = meetingTimeZone.ToTimeZoneInfo();
            }
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
            MeetingTimeZone value = (MeetingTimeZone)propertyBag[this];

            if (value != null)
            {
                value.WriteToXml(writer, this.XmlElementName);
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
            MeetingTimeZone value = propertyBag[this] as MeetingTimeZone;

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
            get { return typeof(MeetingTimeZone); }
        }
    }
}