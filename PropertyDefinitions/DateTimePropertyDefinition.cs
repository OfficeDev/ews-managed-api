// ---------------------------------------------------------------------------
// <copyright file="DateTimePropertyDefinition.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the DateTimePropertyDefinition class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;

    /// <summary>
    /// Represents DateTime property definition.
    /// </summary>
    internal class DateTimePropertyDefinition : PropertyDefinition
    {
        private bool isNullable;

        /// <summary>
        /// Initializes a new instance of the <see cref="DateTimePropertyDefinition"/> class.
        /// </summary>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <param name="uri">The URI.</param>
        /// <param name="version">The version.</param>
        internal DateTimePropertyDefinition(
            string xmlElementName,
            string uri,
            ExchangeVersion version)
            : base(xmlElementName, uri, version)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DateTimePropertyDefinition"/> class.
        /// </summary>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <param name="uri">The URI.</param>
        /// <param name="flags">The flags.</param>
        /// <param name="version">The version.</param>
        internal DateTimePropertyDefinition(
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
        /// Initializes a new instance of the <see cref="DateTimePropertyDefinition"/> class.
        /// </summary>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <param name="uri">The URI.</param>
        /// <param name="flags">The flags.</param>
        /// <param name="version">The version.</param>
        /// <param name="isNullable">Indicates that this property definition is for a nullable property.</param>
        internal DateTimePropertyDefinition(
            string xmlElementName,
            string uri,
            PropertyDefinitionFlags flags,
            ExchangeVersion version,
            bool isNullable)
            : base(
                xmlElementName,
                uri,
                flags,
                version)
        {
            this.isNullable = isNullable;
        }

        /// <summary>
        /// Loads from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <param name="propertyBag">The property bag.</param>
        internal override void LoadPropertyValueFromXml(EwsServiceXmlReader reader, PropertyBag propertyBag)
        {
            string value = reader.ReadElementValue(XmlNamespace.Types, this.XmlElementName);

            propertyBag[this] = reader.Service.ConvertUniversalDateTimeStringToLocalDateTime(value);
        }

        /// <summary>
        /// Loads the property value from json.
        /// </summary>
        /// <param name="value">The JSON value.  Can be a JsonObject, string, number, bool, array, or null.</param>
        /// <param name="service">The service.</param>
        /// <param name="propertyBag">The property bag.</param>
        internal override void LoadPropertyValueFromJson(object value, ExchangeService service, PropertyBag propertyBag)
        {
            string stringValue = value as string;

            if (!String.IsNullOrEmpty(stringValue))
            {
                propertyBag[this] = service.ConvertUniversalDateTimeStringToLocalDateTime(stringValue);
            }
        }

        /// <summary>
        /// Scopes the date time property to the appropriate time zone, if necessary.
        /// </summary>
        /// <param name="service">The service emitting the request.</param>
        /// <param name="dateTime">The date time.</param>
        /// <param name="propertyBag">The property bag.</param>
        /// <param name="isUpdateOperation">Indicates whether the scoping is to be performed in the context of an update operation.</param>
        /// <returns>The converted DateTime.</returns>
        internal virtual DateTime ScopeToTimeZone(
            ExchangeServiceBase service,
            DateTime dateTime,
            PropertyBag propertyBag,
            bool isUpdateOperation)
        {
            try
            {
                DateTime convertedDateTime = EwsUtilities.ConvertTime(
                    dateTime,
                    service.TimeZone,
                    TimeZoneInfo.Utc);

                return new DateTime(convertedDateTime.Ticks, DateTimeKind.Utc);
            }
            catch (TimeZoneConversionException e)
            {
                throw new PropertyException(
                    string.Format(Strings.InvalidDateTime, dateTime),
                    this.Name,
                    e);
            }
        }

        /// <summary>
        /// Writes the property value to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="propertyBag">The property bag.</param>
        /// <param name="isUpdateOperation">Indicates whether the context is an update operation.</param>
        internal override void WritePropertyValueToXml(
            EwsServiceXmlWriter writer,
            PropertyBag propertyBag,
            bool isUpdateOperation)
        {
            object value = propertyBag[this];

            if (value != null)
            {
                writer.WriteStartElement(XmlNamespace.Types, this.XmlElementName);

                DateTime convertedDateTime = GetConvertedDateTime(writer.Service, propertyBag, isUpdateOperation, value);

                writer.WriteValue(EwsUtilities.DateTimeToXSDateTime(convertedDateTime), this.Name);

                writer.WriteEndElement();
            }
        }

        /// <summary>
        /// Writes the json value.
        /// </summary>
        /// <param name="jsonObject">The json object.</param>
        /// <param name="propertyBag">The property bag.</param>
        /// <param name="service">The service.</param>
        /// 
        /// <param name="isUpdateOperation">if set to <c>true</c> [is update operation].</param>
        internal override void WriteJsonValue(JsonObject jsonObject, PropertyBag propertyBag, ExchangeService service, bool isUpdateOperation)
        {
            object value = propertyBag[this];

            if (value != null)
            {
                DateTime convertedDateTime = GetConvertedDateTime(service, propertyBag, isUpdateOperation, value);

                jsonObject.Add(this.XmlElementName, EwsUtilities.DateTimeToXSDateTime(convertedDateTime));
            }
        }

        /// <summary>
        /// Gets the converted date time.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="propertyBag">The property bag.</param>
        /// <param name="isUpdateOperation">if set to <c>true</c> [is update operation].</param>
        /// <param name="value">The value.</param>
        /// <returns></returns>
        private DateTime GetConvertedDateTime(ExchangeServiceBase service, PropertyBag propertyBag, bool isUpdateOperation, object value)
        {
            DateTime dateTime = (DateTime)value;
            DateTime convertedDateTime;

            // If the date/time is unspecified, we may need to scope it to time zone.
            if (dateTime.Kind == DateTimeKind.Unspecified)
            {
                convertedDateTime = this.ScopeToTimeZone(
                    service,
                    (DateTime)value,
                    propertyBag,
                    isUpdateOperation);
            }
            else
            {
                convertedDateTime = dateTime;
            }
            return convertedDateTime;
        }

        /// <summary>
        /// Gets a value indicating whether this property definition is for a nullable type (ref, int?, bool?...).
        /// </summary>
        internal override bool IsNullable
        {
            get { return this.isNullable; }
        }

        /// <summary>
        /// Gets the property type.
        /// </summary>
        public override Type Type
        {
            get { return this.IsNullable ? typeof(DateTime?) : typeof(DateTime); }
        }
    }
}