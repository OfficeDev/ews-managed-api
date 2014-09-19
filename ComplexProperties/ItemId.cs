// ---------------------------------------------------------------------------
// <copyright file="ItemId.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ItemId class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents the Id of an Exchange item.
    /// </summary>
    public class ItemId : ServiceId, IJsonSerializable
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ItemId"/> class.
        /// </summary>
        internal ItemId()
            : base()
        {
        }

        /// <summary>
        /// Defines an implicit conversion between string and ItemId.
        /// </summary>
        /// <param name="uniqueId">The unique Id to convert to ItemId.</param>
        /// <returns>An ItemId initialized with the specified unique Id.</returns>
        public static implicit operator ItemId(string uniqueId)
        {
            return new ItemId(uniqueId);
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.ItemId;
        }

        /// <summary>
        /// Initializes a new instance of ItemId.
        /// </summary>
        /// <param name="uniqueId">The unique Id used to initialize the ItemId.</param>
        public ItemId(string uniqueId)
            : base(uniqueId)
        {
        }
    }
}
