// ---------------------------------------------------------------------------
// <copyright file="GenericItemAttachment.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the GenericItemAttachment enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a strongly typed item attachment.
    /// </summary>
    /// <typeparam name="TItem">Item type.</typeparam>
    public sealed class ItemAttachment<TItem> : ItemAttachment
        where TItem : Item
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ItemAttachment&lt;TItem&gt;"/> class.
        /// </summary>
        /// <param name="owner">The owner of the attachment.</param>
        internal ItemAttachment(Item owner)
            : base(owner)
        {
        }

        /// <summary>
        /// Gets the item associated with the attachment.
        /// </summary>
        public new TItem Item
        {
            get { return (TItem)base.Item; }
            internal set { base.Item = value; }
        }
    }
}
