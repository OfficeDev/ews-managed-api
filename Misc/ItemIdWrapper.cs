// ---------------------------------------------------------------------------
// <copyright file="ItemIdWrapper.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ItemIdWrapper enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents an item Id provided by a ItemId object.
    /// </summary>
    internal class ItemIdWrapper : AbstractItemIdWrapper
    {
        /// <summary>
        /// The ItemId object providing the Id.
        /// </summary>
        private ItemId itemId;

        /// <summary>
        /// Initializes a new instance of ItemIdWrapper.
        /// </summary>
        /// <param name="itemId">The ItemId object providing the Id.</param>
        internal ItemIdWrapper(ItemId itemId)
        {
            EwsUtilities.Assert(
                itemId != null,
                "ItemIdWrapper.ctor",
                "itemId is null");

            this.itemId = itemId;
        }

        /// <summary>
        /// Writes the Id encapsulated in the wrapper to XML.
        /// </summary>
        /// <param name="writer">The writer to write the Id to.</param>
        internal override void WriteToXml(EwsServiceXmlWriter writer)
        {
            this.itemId.WriteToXml(writer);
        }

        /// <summary>
        /// Creates a JSON representation of this object.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <returns>
        /// A Json value (either a JsonObject, an array of Json values, or a Json primitive)
        /// </returns>
        internal override object IternalToJson(ExchangeService service)
        {
            return this.itemId.InternalToJson(service);
        }
    }
}
