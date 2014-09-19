// ---------------------------------------------------------------------------
// <copyright file="ItemChange.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ItemChange class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a change on an item as returned by a synchronization operation.
    /// </summary>
    public sealed class ItemChange : Change
    {
        private bool isRead;

        /// <summary>
        /// Initializes a new instance of ItemChange.
        /// </summary>
        internal ItemChange()
            : base()
        {
        }

        /// <summary>
        /// Creates an ItemId instance.
        /// </summary>
        /// <returns>A ItemId.</returns>
        internal override ServiceId CreateId()
        {
            return new ItemId();
        }

        /// <summary>
        /// Gets the item the change applies to. Item is null when ChangeType is equal to
        /// either ChangeType.Delete or ChangeType.ReadFlagChange. In those cases, use the
        /// ItemId property to retrieve the Id of the item that was deleted or whose IsRead
        /// property changed.
        /// </summary>
        public Item Item
        {
            get { return (Item)this.ServiceObject; }
        }

        /// <summary>
        /// Gets the IsRead property for the item that the change applies to. IsRead is
        /// only valid when ChangeType is equal to ChangeType.ReadFlagChange.
        /// </summary>
        public bool IsRead
        {
            get { return this.isRead; }
            internal set { this.isRead = value; }
        }

        /// <summary>
        /// Gets the Id of the item the change applies to.
        /// </summary>
        public ItemId ItemId
        {
            get { return (ItemId)this.Id; }
        }
    }
}
