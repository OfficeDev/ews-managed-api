/*
 * Exchange Web Services Managed API
 *
 * Copyright (c) Microsoft Corporation
 * All rights reserved.
 *
 * MIT License
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this
 * software and associated documentation files (the "Software"), to deal in the Software
 * without restriction, including without limitation the rights to use, copy, modify, merge,
 * publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
 * to whom the Software is furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all copies or
 * substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
 * INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
 * PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
 * FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
 * OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
 * DEALINGS IN THE SOFTWARE.
 */

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