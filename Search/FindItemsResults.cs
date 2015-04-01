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
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Text;

    /// <summary>
    /// Represents the results of an item search operation.
    /// </summary>
    /// <typeparam name="TItem">The type of item returned by the search operation.</typeparam>
    public sealed class FindItemsResults<TItem> : IEnumerable<TItem>
        where TItem : Item
    {
        private int totalCount;
        private int? nextPageOffset;
        private bool moreAvailable;
        private Collection<TItem> items = new Collection<TItem>();
        private Collection<HighlightTerm> highlightTerms = new Collection<HighlightTerm>();

        /// <summary>
        /// Initializes a new instance of the <see cref="FindItemsResults&lt;T&gt;"/> class.
        /// </summary>
        internal FindItemsResults()
        {
        }

        /// <summary>
        /// Gets the total number of items matching the search criteria available in the searched folder.
        /// </summary>
        public int TotalCount
        {
            get { return this.totalCount; }
            internal set { this.totalCount = value; }
        }

        /// <summary>
        /// Gets the offset that should be used with ItemView to retrieve the next page of items in a FindItems operation.
        /// </summary>
        public int? NextPageOffset
        {
            get { return this.nextPageOffset; }
            internal set { this.nextPageOffset = value; }
        }

        /// <summary>
        /// Gets a value indicating whether more items matching the search criteria
        /// are available in the searched folder. 
        /// </summary>
        public bool MoreAvailable
        {
            get { return this.moreAvailable; }
            internal set { this.moreAvailable = value; }
        }

        /// <summary>
        /// Gets a collection containing the items that were found by the search operation.
        /// </summary>
        public Collection<TItem> Items
        {
            get { return this.items; }
        }

        /// <summary>
        /// Gets a collection containing the highlight terms that were found by the search operation.
        /// </summary>
        public Collection<HighlightTerm> HighlightTerms
        {
            get { return this.highlightTerms; }
        }

        #region IEnumerable<T> Members

        /// <summary>
        /// Returns an enumerator that iterates through the collection.
        /// </summary>
        /// <returns>
        /// A <see cref="T:System.Collections.Generic.IEnumerator`1"/> that can be used to iterate through the collection.
        /// </returns>
        public IEnumerator<TItem> GetEnumerator()
        {
            return this.items.GetEnumerator();
        }

        #endregion

        #region IEnumerable Members

        /// <summary>
        /// Returns an enumerator that iterates through a collection.
        /// </summary>
        /// <returns>
        /// An <see cref="T:System.Collections.IEnumerator"/> object that can be used to iterate through the collection.
        /// </returns>
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.items.GetEnumerator();
        }

        #endregion
    }
}