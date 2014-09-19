// ---------------------------------------------------------------------------
// <copyright file="GenericPropertyDefinition.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the GenericPropertyDefinition class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;

    /// <summary>
    /// Represents generic property definition.
    /// </summary>
    /// <typeparam name="TPropertyValue">Property value type. Constrained to be a value type.</typeparam>
    internal class GenericPropertyDefinition<TPropertyValue> : TypedPropertyDefinition where TPropertyValue : struct
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="GenericPropertyDefinition&lt;T&gt;"/> class.
        /// </summary>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <param name="uri">The URI.</param>
        /// <param name="version">The version.</param>
        internal GenericPropertyDefinition(
            string xmlElementName,
            string uri,
            ExchangeVersion version)
            : base(
                xmlElementName,
                uri,
                version)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="GenericPropertyDefinition&lt;T&gt;"/> class.
        /// </summary>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <param name="uri">The URI.</param>
        /// <param name="flags">The flags.</param>
        /// <param name="version">The version.</param>
        internal GenericPropertyDefinition(
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
        /// Initializes a new instance of the <see cref="GenericPropertyDefinition&lt;T&gt;"/> class.
        /// </summary>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <param name="uri">The URI.</param>
        /// <param name="flags">The flags.</param>
        /// <param name="version">The version.</param>
        /// <param name="isNullable">if set to true, property value is nullable.</param>
        internal GenericPropertyDefinition(
            string xmlElementName,
            string uri,
            PropertyDefinitionFlags flags,
            ExchangeVersion version,
            bool isNullable)
            : base(
                xmlElementName,
                uri,
                flags,
                version,
                isNullable)
        {
        }

        /// <summary>
        /// Parses the specified value.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <returns>Value of string.</returns>
        internal override object Parse(string value)
        {
            return EwsUtilities.Parse<TPropertyValue>(value);
        }

        /// <summary>
        /// Gets the property type.
        /// </summary>
        public override Type Type
        {
            get { return this.IsNullable ? typeof(Nullable<TPropertyValue>) : typeof(TPropertyValue); }
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
            jsonObject.Add(this.XmlElementName, propertyBag[this]);
        }
    }
}
