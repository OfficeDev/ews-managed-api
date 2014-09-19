// ---------------------------------------------------------------------------
// <copyright file="ByteArrayPropertyDefinition.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ByteArrayPropertyDefinition class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;

    /// <summary>
    /// Represents byte array property definition.
    /// </summary>
    internal sealed class ByteArrayPropertyDefinition : TypedPropertyDefinition
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ByteArrayPropertyDefinition"/> class.
        /// </summary>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <param name="uri">The URI.</param>
        /// <param name="flags">The flags.</param>
        /// <param name="version">The version.</param>
        internal ByteArrayPropertyDefinition(
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
        /// Parses the specified value.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <returns>Byte array value.</returns>
        internal override object Parse(string value)
        {
            return Convert.FromBase64String(value);
        }

        /// <summary>
        /// Converts byte array property to a string.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <returns>Byte array value.</returns>
        internal override string ToString(object value)
        {
            return Convert.ToBase64String((byte[])value);
        }

        /// <summary>
        /// Gets a value indicating whether this property definition is for a nullable type (ref, int?, bool?...).
        /// </summary>
        internal override bool IsNullable
        {
            get { return true; }
        }

        /// <summary>
        /// Gets the property type.
        /// </summary>
        public override Type Type
        {
            get { return typeof(byte[]); }
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
            if (propertyBag[this] != null)
            {
                jsonObject.Add(this.XmlElementName, this.ToString(propertyBag[this]));
            }
        }
    }
}
