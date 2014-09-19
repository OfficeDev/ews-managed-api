// ---------------------------------------------------------------------------
// <copyright file="TypedPropertyDefinition.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the TypedPropertyDefinition class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Represents typed property definition.
    /// </summary>
    internal abstract class TypedPropertyDefinition : PropertyDefinition
    {
        private bool isNullable;

        /// <summary>
        /// Initializes a new instance of the <see cref="TypedPropertyDefinition"/> class.
        /// </summary>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <param name="uri">The URI.</param>
        /// <param name="version">The version.</param>
        internal TypedPropertyDefinition(
            string xmlElementName,
            string uri,
            ExchangeVersion version)
            : base(
                xmlElementName,
                uri,
                version)
        {
            this.isNullable = false;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="TypedPropertyDefinition"/> class.
        /// </summary>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <param name="uri">The URI.</param>
        /// <param name="flags">The flags.</param>
        /// <param name="version">The version.</param>
        internal TypedPropertyDefinition(
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
        /// Initializes a new instance of the <see cref="TypedPropertyDefinition"/> class.
        /// </summary>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <param name="uri">The URI.</param>
        /// <param name="flags">The flags.</param>
        /// <param name="version">The version.</param>
        /// <param name="isNullable">Indicates that this property definition is for a nullable property.</param>
        internal TypedPropertyDefinition(
            string xmlElementName,
            string uri,
            PropertyDefinitionFlags flags,
            ExchangeVersion version,
            bool isNullable)
            : this(
                xmlElementName,
                uri,
                flags,
                version)
        {
            this.isNullable = isNullable;
        }

        /// <summary>
        /// Parses the specified value.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <returns>Typed value.</returns>
        internal abstract object Parse(string value);

        /// <summary>
        /// Gets a value indicating whether this property definition is for a nullable type (ref, int?, bool?...).
        /// </summary>
        internal override bool IsNullable
        {
            get { return this.isNullable; }
        }

        /// <summary>
        /// Convert instance to string.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <returns>String representation of property value.</returns>
        internal virtual string ToString(object value)
        {
            return value.ToString();
        }

        /// <summary>
        /// Loads from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <param name="propertyBag">The property bag.</param>
        internal override void LoadPropertyValueFromXml(EwsServiceXmlReader reader, PropertyBag propertyBag)
        {
            string value = reader.ReadElementValue(XmlNamespace.Types, this.XmlElementName);

            if (!string.IsNullOrEmpty(value))
            {
                propertyBag[this] = this.Parse(value);
            }
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

            if (!string.IsNullOrEmpty(stringValue))
            {
                propertyBag[this] = this.Parse(stringValue);
            }
            else if (value != null)
            {
                propertyBag[this] = this.Parse(value.ToString());
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
            object value = propertyBag[this];

            if (value != null)
            {
                writer.WriteElementValue(XmlNamespace.Types, this.XmlElementName, this.Name, value);
            }
        }
    }
}