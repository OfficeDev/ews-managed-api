// ---------------------------------------------------------------------------
// <copyright file="ICustomXmlUpdateSerializer.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ICustomXmlUpdateSerializer interface.</summary>
//-----------------------------------------------------------------------

using System.Collections.Generic;

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Interface defined for properties that produce their own update serialization.
    /// </summary>
    internal interface ICustomUpdateSerializer
    {
        /// <summary>
        /// Writes the update to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="ewsObject">The ews object.</param>
        /// <param name="propertyDefinition">Property definition.</param>
        /// <returns>True if property generated serialization.</returns>
        bool WriteSetUpdateToXml(
            EwsServiceXmlWriter writer,
            ServiceObject ewsObject,
            PropertyDefinition propertyDefinition);

        /// <summary>
        /// Writes the deletion update to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="ewsObject">The ews object.</param>
        /// <returns>True if property generated serialization.</returns>
        bool WriteDeleteUpdateToXml(EwsServiceXmlWriter writer, ServiceObject ewsObject);

        /// <summary>
        /// Writes the update to Json.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="ewsObject">The ews object.</param>
        /// <param name="propertyDefinition">Property definition.</param>
        /// <param name="updates">The updates.</param>
        /// <returns>
        /// True if property generated serialization.
        /// </returns>
        bool WriteSetUpdateToJson(
            ExchangeService service,
            ServiceObject ewsObject,
            PropertyDefinition propertyDefinition,
            List<JsonObject> updates);

        /// <summary>
        /// Writes the deletion update to Json.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="ewsObject">The ews object.</param>
        /// <param name="updates">The updates.</param>
        /// <returns>
        /// True if property generated serialization.
        /// </returns>
        bool WriteDeleteUpdateToJson(ExchangeService service, ServiceObject ewsObject, List<JsonObject> updates);
    }
}
