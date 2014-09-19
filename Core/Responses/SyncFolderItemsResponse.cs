// ---------------------------------------------------------------------------
// <copyright file="SyncFolderItemsResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the SyncFolderItemsResponse class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents the response to a folder items synchronization operation.
    /// </summary>
    public sealed class SyncFolderItemsResponse : SyncResponse<Item, ItemChange>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="SyncFolderItemsResponse"/> class.
        /// </summary>
        /// <param name="propertySet">PropertySet from request.</param>
        internal SyncFolderItemsResponse(PropertySet propertySet)
            : base(propertySet)
        {
        }

        /// <summary>
        /// Gets the name of the includes last in range XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetIncludesLastInRangeXmlElementName()
        {
            return XmlElementNames.IncludesLastItemInRange;
        }

        /// <summary>
        /// Creates an item change instance.
        /// </summary>
        /// <returns>ItemChange instance</returns>
        internal override ItemChange CreateChangeInstance()
        {
            return new ItemChange();
        }

        /// <summary>
        /// Gets the name of the change element.
        /// </summary>
        /// <returns>Change element name.</returns>
        internal override string GetChangeElementName()
        {
            return XmlElementNames.Item;
        }

        /// <summary>
        /// Gets the name of the change id element.
        /// </summary>
        /// <returns>Change id element name.</returns>
        internal override string GetChangeIdElementName()
        {
            return XmlElementNames.ItemId;
        }

        /// <summary>
        /// Gets a value indicating whether this request returns full or summary properties.
        /// </summary>
        /// <value>
        /// <c>true</c> if summary properties only; otherwise, <c>false</c>.
        /// </value>
        internal override bool SummaryPropertiesOnly
        {
            get { return true; }
        }
    }
}
