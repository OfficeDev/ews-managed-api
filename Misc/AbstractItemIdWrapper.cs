// ---------------------------------------------------------------------------
// <copyright file="AbstractItemIdWrapper.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the AbstractItemIdWrapper enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents the abstraction of an item Id.
    /// </summary>
    internal abstract class AbstractItemIdWrapper : IJsonSerializable
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="AbstractItemIdWrapper"/> class.
        /// </summary>
        internal AbstractItemIdWrapper()
        {
        }

        /// <summary>
        /// Obtains the ItemBase object associated with the wrapper.
        /// </summary>
        /// <returns>The ItemBase object associated with the wrapper.</returns>
        public virtual Item GetItem()
        {
            return null;
        }

        /// <summary>
        /// Writes the Id encapsulated in the wrapper to XML.
        /// </summary>
        /// <param name="writer">The writer to write the Id to.</param>
        internal abstract void WriteToXml(EwsServiceXmlWriter writer);

        /// <summary>
        /// Creates a JSON representation of this object.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <returns>
        /// A Json value (either a JsonObject, an array of Json values, or a Json primitive)
        /// </returns>
        object IJsonSerializable.ToJson(ExchangeService service)
        {
            return this.IternalToJson(service);
        }

        /// <summary>
        /// Creates a JSON representation of this object.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <returns>
        /// A Json value (either a JsonObject, an array of Json values, or a Json primitive)
        /// </returns>
        internal abstract object IternalToJson(ExchangeService service);
    }
}
