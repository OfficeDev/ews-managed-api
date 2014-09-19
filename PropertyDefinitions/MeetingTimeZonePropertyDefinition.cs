// ---------------------------------------------------------------------------
// <copyright file="MeetingTimeZonePropertyDefinition.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the MeetingTimeZonePropertyDefinition class.</summary>
//-----------------------------------------------------------------------

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