// ---------------------------------------------------------------------------
// <copyright file="TimeSpanPropertyDefinition.cs" company="Microsoft">
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
    /// Represents TimeSpan property definition.
    /// </summary>
    internal class TimeSpanPropertyDefinition : GenericPropertyDefinition<TimeSpan>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="TimeSpanPropertyDefinition"/> class.
        /// </summary>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <param name="uri">The URI.</param>
        /// <param name="flags">The flags.</param>
        /// <param name="version">The version.</param>
        internal TimeSpanPropertyDefinition(
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
        /// <returns>TimeSpan value.</returns>
        internal override object Parse(string value)
        {
            return EwsUtilities.XSDurationToTimeSpan(value);
        }

        /// <summary>
        /// Converts instance to a string.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <returns>TimeSpan value.</returns>
        internal override string ToString(object value)
        {
            return EwsUtilities.TimeSpanToXSDuration((TimeSpan)value);
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
