// ---------------------------------------------------------------------------
// <copyright file="ItemCollection.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ItemCollection class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Text;
    using System.Xml;

    /// <summary>
    /// Represents a collection of items.
    /// </summary>
    /// <typeparam name="TItem">The type of item the collection contains.</typeparam>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public sealed class ItemCollection<TItem> : ComplexProperty, IEnumerable<TItem>, IJsonCollectionDeserializer
        where TItem : Item
    {
        private List<TItem> items = new List<TItem>();

        /// <summary>
        /// Initializes a new instance of the <see cref="ItemCollection&lt;TItem&gt;"/> class.
        /// </summary>
        internal ItemCollection()
            : base()
        {
        }

        /// <summary>
        /// Loads from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <param name="localElementName">Name of the local element.</param>
        internal override void LoadFromXml(EwsServiceXmlReader reader, string localElementName)
        {
            reader.EnsureCurrentNodeIsStartElement(XmlNamespace.Types, localElementName);
            if (!reader.IsEmptyElement)
            {
                do
                {
                    reader.Read();

                    if (reader.NodeType == XmlNodeType.Element)
                    {
                        TItem item = EwsUtilities.CreateEwsObjectFromXmlElementName<Item>(
                            reader.Service,
                            reader.LocalName) as TItem;

                        if (item == null)
                        {
                            reader.SkipCurrentElement();
                        }
                        else
                        {
                            item.LoadFromXml(reader, true /* clearPropertyBag */);

                            this.items.Add(item);
                        }
                    }
                }
                while (!reader.IsEndElement(XmlNamespace.Types, localElementName));
            }
        }

        /// <summary>
        /// Loads from json collection.
        /// </summary>
        /// <param name="jsonCollection">The json collection.</param>
        /// <param name="service">The service.</param>
        void IJsonCollectionDeserializer.CreateFromJsonCollection(object[] jsonCollection, ExchangeService service)
        {
            foreach (object entry in jsonCollection)
            {
                JsonObject jsonServiceObject = entry as JsonObject;

                TItem item = EwsUtilities.CreateEwsObjectFromXmlElementName<Item>(
                    service,
                    jsonServiceObject.ReadTypeString()) as TItem;

                item.LoadFromJson(jsonServiceObject, service, true);

                this.items.Add(item);
            }
        }

        /// <summary>
        /// Loads from json collection to update the existing collection element.
        /// </summary>
        /// <param name="jsonCollection">The json collection.</param>
        /// <param name="service">The service.</param>
        void IJsonCollectionDeserializer.UpdateFromJsonCollection(object[] jsonCollection, ExchangeService service)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Gets the total number of items in the collection.
        /// </summary>
        public int Count
        {
            get { return this.items.Count; }
        }

        /// <summary>
        /// Gets the item at the specified index.
        /// </summary>
        /// <param name="index">The zero-based index of the item to get.</param>
        /// <returns>The item at the specified index.</returns>
        public TItem this[int index]
        {
            get
            {
                if (index < 0 || index >= this.Count)
                {
                    throw new ArgumentOutOfRangeException("index", Strings.IndexIsOutOfRange);
                }

                return this.items[index];
            }
        }

        #region IEnumerable<TItem> Members

        /// <summary>
        /// Gets an enumerator that iterates through the elements of the collection.
        /// </summary>
        /// <returns>An IEnumerator for the collection.</returns>
        public IEnumerator<TItem> GetEnumerator()
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
