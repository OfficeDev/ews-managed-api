// ---------------------------------------------------------------------------
// <copyright file="ItemGroup.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ItemGroup class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System.Collections.Generic;
    using System.Collections.ObjectModel;

    /// <summary>
    /// Represents a group of items as returned by grouped item search operations.
    /// </summary>
    /// <typeparam name="TItem">The type of item in the group.</typeparam>
    public sealed class ItemGroup<TItem>
        where TItem : Item
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ItemGroup&lt;TItem&gt;"/> class.
        /// </summary>
        /// <param name="groupIndex">Index of the group.</param>
        /// <param name="items">The items.</param>
        internal ItemGroup(string groupIndex, IList<TItem> items)
        {
            EwsUtilities.Assert(
                items != null,
                "ItemGroup.ctor",
                "items is null");

            this.GroupIndex = groupIndex;
            this.Items = new Collection<TItem>(items);
        }

        /// <summary>
        /// Gets an index identifying the group.
        /// </summary>
        public string GroupIndex
        {
            get; private set;
        }

        /// <summary>
        /// Gets a collection of the items in this group.
        /// </summary>
        public Collection<TItem> Items
        {
            get; private set;
        }
    }
}