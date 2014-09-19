// ---------------------------------------------------------------------------
// <copyright file="TimeZonePropertyDefinition.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the TimeSpanPropertyDefinition class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;

    /// <summary>
    /// Represents a property definition for properties of type TimeZoneInfo.
    /// </summary>
    internal class TimeZonePropertyDefinition : PropertyDefinition
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="TimeZonePropertyDefinition"/> class.
        /// </summary>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <param name="uri">The URI.</param>
        /// <param name="flags">The flags.</param>
        /// <param name="version">The version.</param>
        internal TimeZonePropertyDefinition(
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
            TimeZoneDefinition timeZoneDefinition = new TimeZoneDefinition();
            timeZoneDefinition.LoadFromXml(reader, this.XmlElementName);

            propertyBag[this] = timeZoneDefinition.ToTimeZoneInfo();
        }

        /// <summary>
        /// Loads the property value from json.
        /// </summary>
        /// <param name="value">The JSON value.  Can be a JsonObject, string, number, bool, array, or null.</param>
        /// <param name="service">The service.</param>
        /// <param name="propertyBag">The property bag.</param>
        internal override void LoadPropertyValueFromJson(object value, ExchangeService service, PropertyBag propertyBag)
        {
            TimeZoneDefinition timeZoneDefinition = new TimeZoneDefinition();
            JsonObject jsonTimeZoneProperty = value as JsonObject;

            if (jsonTimeZoneProperty != null)
            {
                timeZoneDefinition.LoadFromJson(jsonTimeZoneProperty, service);
            }

            propertyBag[this] = timeZoneDefinition.ToTimeZoneInfo();
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
            TimeZoneInfo value = (TimeZoneInfo)propertyBag[this];

            if (value != null)
            {
                // We emit time zone properties only if we have not emitted the time zone SOAP header
                // or if this time zone is different from that of the service through which the request
                // is being emitted.
                if (!writer.IsTimeZoneHeaderEmitted || value != writer.Service.TimeZone)
                {
                    TimeZoneDefinition timeZoneDefinition = new TimeZoneDefinition(value);

                    timeZoneDefinition.WriteToXml(writer, this.XmlElementName);
                }
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
            TimeZoneInfo value = propertyBag[this] as TimeZoneInfo;

            if (value != null &&
                value != service.TimeZone)
            {
                TimeZoneDefinition timeZoneDefinition = new TimeZoneDefinition(value);

                jsonObject.Add(this.XmlElementName, timeZoneDefinition.InternalToJson(service));
            }
        }

        /// <summary>
        /// Gets the property type.
        /// </summary>
        public override Type Type
        {
            get { return typeof(TimeZoneInfo); }
        }
    }
}