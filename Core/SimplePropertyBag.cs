#region License

// Exchange Web Services Managed API
// 
// Copyright (c) Microsoft Corporation
// All rights reserved. 
//
// MIT License
//
// Permission is hereby granted, free of charge, to any person obtaining a copy of this
// software and associated documentation files (the "Software"), to deal in the Software
// without restriction, including without limitation the rights to use, copy, modify, merge,
// publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
// to whom the Software is furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all copies or
// substantial portions of the Software.

// THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
// INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
// PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
// FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
// OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
// DEALINGS IN THE SOFTWARE.

#endregion

//-----------------------------------------------------------------------
// <summary>Defines the SimplePropertyBag class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System.Collections.Generic;

    /// <summary>
    /// Represents a simple property bag.
    /// </summary>
    /// <typeparam name="TKey">The type of the key.</typeparam>
    internal class SimplePropertyBag<TKey> : IEnumerable<KeyValuePair<TKey, object>>
    {
        private Dictionary<TKey, object> items = new Dictionary<TKey, object>();
        private List<TKey> removedItems = new List<TKey>();
        private List<TKey> addedItems = new List<TKey>();
        private List<TKey> modifiedItems = new List<TKey>();

        /// <summary>
        /// Add item to change list.
        /// </summary>
        /// <param name="key">The key.</param>
        /// <param name="changeList">The change list.</param>
        private static void InternalAddItemToChangeList(TKey key, List<TKey> changeList)
        {
            if (!changeList.Contains(key))
            {
                changeList.Add(key);
            }
        }

        /// <summary>
        /// Triggers dispatch of the change event.
        /// </summary>
        private void Changed()
        {
            if (this.OnChange != null)
            {
                this.OnChange();
            }
        }

        /// <summary>
        /// Remove item.
        /// </summary>
        /// <param name="key">The key.</param>
        private void InternalRemoveItem(TKey key)
        {
            object value;

            if (this.TryGetValue(key, out value))
            {
                this.items.Remove(key);
                this.removedItems.Add(key);
                this.Changed();
            }
        }

        /// <summary>
        /// Gets the added items.
        /// </summary>
        /// <value>The added items.</value>
        internal IEnumerable<TKey> AddedItems
        {
            get { return this.addedItems; }
        }

        /// <summary>
        /// Gets the removed items.
        /// </summary>
        /// <value>The removed items.</value>
        internal IEnumerable<TKey> RemovedItems
        {
            get { return this.removedItems; }
        }

        /// <summary>
        /// Gets the modified items.
        /// </summary>
        /// <value>The modified items.</value>
        internal IEnumerable<TKey> ModifiedItems
        {
            get { return this.modifiedItems; }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="SimplePropertyBag&lt;TKey&gt;"/> class.
        /// </summary>
        public SimplePropertyBag()
        {
        }

        /// <summary>
        /// Clears the change log.
        /// </summary>
        public void ClearChangeLog()
        {
            this.removedItems.Clear();
            this.addedItems.Clear();
            this.modifiedItems.Clear();
        }

        /// <summary>
        /// Determines whether the specified key is in the property bag.
        /// </summary>
        /// <param name="key">The key.</param>
        /// <returns>
        ///     <c>true</c> if the specified key exists; otherwise, <c>false</c>.
        /// </returns>
        public bool ContainsKey(TKey key)
        {
            return this.items.ContainsKey(key);
        }

        /// <summary>
        /// Tries to get value.
        /// </summary>
        /// <param name="key">The key.</param>
        /// <param name="value">The value.</param>
        /// <returns>True if value exists in property bag.</returns>
        public bool TryGetValue(TKey key, out object value)
        {
            return this.items.TryGetValue(key, out value);
        }

        /// <summary>
        /// Gets or sets the <see cref="System.Object"/> with the specified key.
        /// </summary>
        /// <param name="key">Key.</param>
        /// <value>Value associated with key.</value>
        public object this[TKey key]
        {
            get
            {
                object value;

                if (this.TryGetValue(key, out value))
                {
                    return value;
                }
                else
                {
                    return null;
                }
            }

            set
            {
                if (value == null)
                {
                    this.InternalRemoveItem(key);
                }
                else
                {
                    // If the item was to be deleted, the deletion becomes an update.
                    if (this.removedItems.Remove(key))
                    {
                        InternalAddItemToChangeList(key, this.modifiedItems);
                    }
                    else
                    {
                        // If the property value was not set, we have a newly set property.
                        if (!this.ContainsKey(key))
                        {
                            InternalAddItemToChangeList(key, this.addedItems);
                        }
                        else
                        {
                            // The last case is that we have a modified property.
                            if (!this.modifiedItems.Contains(key))
                            {
                                InternalAddItemToChangeList(key, this.modifiedItems);
                            }
                        }
                    }

                    this.items[key] = value;
                    this.Changed();
                }
            }
        }

        /// <summary>
        /// Occurs when Changed.
        /// </summary>
        public event PropertyBagChangedDelegate OnChange;

        #region IEnumerable<KeyValuePair<TKey,object>> Members

        /// <summary>
        /// Gets an enumerator that iterates through the elements of the collection.
        /// </summary>
        /// <returns>An IEnumerator for the collection.</returns>
        public IEnumerator<KeyValuePair<TKey, object>> GetEnumerator()
        {
            return this.items.GetEnumerator();
        }

        #endregion

        #region IEnumerable Members

        /// <summary>
        /// Gets an enumerator that iterates through the elements of the collection.
        /// </summary>
        /// <returns>An IEnumerator for the collection.</returns>
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.items.GetEnumerator();
        }

        #endregion
    }
}
