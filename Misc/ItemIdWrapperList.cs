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
// <summary>Defines the ItemIdWrapperList class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a list a abstracted item Ids.
    /// </summary>
    internal class ItemIdWrapperList : IEnumerable<AbstractItemIdWrapper>
    {
        /// <summary>
        /// List of <see cref="Microsoft.Exchange.WebServices.Data.Item"/>.
        /// </summary>
        private List<AbstractItemIdWrapper> itemIds = new List<AbstractItemIdWrapper>();

        /// <summary>
        /// Initializes a new instance of the <see cref="ItemIdWrapperList"/> class.
        /// </summary>
        internal ItemIdWrapperList()
        {
        }

        /// <summary>
        /// Adds the specified item.
        /// </summary>
        /// <param name="item">The item.</param>
        internal void Add(Item item)
        {
            this.itemIds.Add(new ItemWrapper(item));
        }

        /// <summary>
        /// Adds the range.
        /// </summary>
        /// <param name="items">The items.</param>
        internal void AddRange(IEnumerable<Item> items)
        {
            foreach (Item item in items)
            {
                this.Add(item);
            }
        }

        /// <summary>
        /// Adds the specified item id.
        /// </summary>
        /// <param name="itemId">The item id.</param>
        internal void Add(ItemId itemId)
        {
            this.itemIds.Add(new ItemIdWrapper(itemId));
        }

        /// <summary>
        /// Adds the range.
        /// </summary>
        /// <param name="itemIds">The item ids.</param>
        internal void AddRange(IEnumerable<ItemId> itemIds)
        {
            foreach (ItemId itemId in itemIds)
            {
                this.Add(itemId);
            }
        }

        /// <summary>
        /// Writes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="ewsNamesapce">The ews namesapce.</param>
        /// <param name="xmlElementName">Name of the XML element.</param>
        internal void WriteToXml(
            EwsServiceXmlWriter writer,
            XmlNamespace ewsNamesapce,
            string xmlElementName)
        {
            if (this.Count > 0)
            {
                writer.WriteStartElement(ewsNamesapce, xmlElementName);

                foreach (AbstractItemIdWrapper itemIdWrapper in this.itemIds)
                {
                    itemIdWrapper.WriteToXml(writer);
                }

                writer.WriteEndElement();
            }
        }

        /// <summary>
        /// Serializes the property to a Json value.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <returns></returns>
        internal object InternalToJson(ExchangeService service)
        {
            List<object> jsonArray = new List<object>();

            foreach (AbstractItemIdWrapper itemIdWraper in this.itemIds)
            {
                jsonArray.Add(((IJsonSerializable)itemIdWraper).ToJson(service));
            }

            return jsonArray.ToArray();
        }

        /// <summary>
        /// Gets the count.
        /// </summary>
        /// <value>The count.</value>
        internal int Count
        {
            get { return this.itemIds.Count; }
        }

        /// <summary>
        /// Gets the <see cref="Microsoft.Exchange.WebServices.Data.Item"/> at the specified index.
        /// </summary>
        /// <param name="index">the index</param>
        internal Item this[int index]
        {
            get { return this.itemIds[index].GetItem(); }
        }

        #region IEnumerable<AbstractItemIdWrapper> Members

        /// <summary>
        /// Gets an enumerator that iterates through the elements of the collection.
        /// </summary>
        /// <returns>An IEnumerator for the collection.</returns>
        public IEnumerator<AbstractItemIdWrapper> GetEnumerator()
        {
            return this.itemIds.GetEnumerator();
        }

        #endregion

        #region IEnumerable Members

        /// <summary>
        /// Gets an enumerator that iterates through the elements of the collection.
        /// </summary>
        /// <returns>An IEnumerator for the collection.</returns>
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.itemIds.GetEnumerator();
        }

        #endregion
    }
}
