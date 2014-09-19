// ---------------------------------------------------------------------------
// <copyright file="ItemWrapper.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ItemWrapper enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents an item Id provided by a ItemBase object.
    /// </summary>
    internal class ItemWrapper : AbstractItemIdWrapper
    {
        /// <summary>
        /// The ItemBase object providing the Id.
        /// </summary>
        private Item item;

        /// <summary>
        /// Initializes a new instance of ItemWrapper.
        /// </summary>
        /// <param name="item">The ItemBase object provinding the Id.</param>
        internal ItemWrapper(Item item)
        {
            EwsUtilities.Assert(
                item != null,
                "ItemWrapper.ctor",
                "item is null");
            EwsUtilities.Assert(
                !item.IsNew,
                "ItemWrapper.ctor",
                "item does not have an Id");

            this.item = item;
        }

        /// <summary>
        /// Obtains the ItemBase object associated with the wrapper.
        /// </summary>
        /// <returns>The ItemBase object associated with the wrapper.</returns>
        public override Item GetItem()
        {
            return this.item;
        }

        /// <summary>
        /// Writes the Id encapsulated in the wrapper to XML.
        /// </summary>
        /// <param name="writer">The writer to write the Id to.</param>
        internal override void WriteToXml(EwsServiceXmlWriter writer)
        {
            this.item.Id.WriteToXml(writer);
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
            return ((IJsonSerializable)this.item.Id).ToJson(service);
        }
    }
}
