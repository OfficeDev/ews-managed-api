// ---------------------------------------------------------------------------
// <copyright file="ListValuePropertyDefinition.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ListValuePropertyDefinition class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;

    /// <summary>
    /// Represents property definition for type represented by xs:list of values in schema.
    /// </summary>
    /// <typeparam name="TPropertyValue">Property value type. Constrained to be a value type.</typeparam>
    internal class ListValuePropertyDefinition<TPropertyValue> : GenericPropertyDefinition<TPropertyValue> where TPropertyValue : struct
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ListValuePropertyDefinition&lt;TPropertyValue&gt;"/> class.
        /// </summary>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <param name="uri">The URI.</param>
        /// <param name="flags">The flags.</param>
        /// <param name="version">The version.</param>
        internal ListValuePropertyDefinition(
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
        /// <returns>Value of string.</returns>
        internal override object Parse(string value)
        {
            // xs:list values are sent as a space-separated list; convert to comma-separated for EwsUtilities.Parse.
            string commaSeparatedValue = string.IsNullOrEmpty(value) ? value : value.Replace(' ', ',');
            return EwsUtilities.Parse<TPropertyValue>(commaSeparatedValue);
        }
    }
}
