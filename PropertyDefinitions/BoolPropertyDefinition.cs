// ---------------------------------------------------------------------------
// <copyright file="BoolPropertyDefinition.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the BoolPropertyDefinition class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;

    /// <summary>
    /// Represents Boolean property definition
    /// </summary>
    internal sealed class BoolPropertyDefinition : GenericPropertyDefinition<bool>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="BoolPropertyDefinition"/> class.
        /// </summary>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <param name="uri">The URI.</param>
        /// <param name="version">The version.</param>
        internal BoolPropertyDefinition(
            string xmlElementName,
            string uri,
            ExchangeVersion version)
            : base(xmlElementName, uri, version)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="BoolPropertyDefinition"/> class.
        /// </summary>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <param name="uri">The URI.</param>
        /// <param name="flags">The flags.</param>
        /// <param name="version">The version.</param>
        internal BoolPropertyDefinition(
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
        /// Initializes a new instance of the <see cref="BoolPropertyDefinition"/> class.
        /// </summary>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <param name="uri">The URI.</param>
        /// <param name="flags">The flags.</param>
        /// <param name="version">The version.</param>
        /// <param name="isNullable">Indicates that this property definition is for a nullable property.</param>
        internal BoolPropertyDefinition(
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
        /// Convert instance to string.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <returns>String representation of Boolean property.</returns>
        internal override string ToString(object value)
        {
            return EwsUtilities.BoolToXSBool((bool)value);
        }
    }
}
