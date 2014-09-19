// ---------------------------------------------------------------------------
// <copyright file="StringList.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the StringList class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a list of strings.
    /// </summary>
    public sealed class StringList : ComplexProperty, IEnumerable<string>, IJsonCollectionDeserializer
    {
        private List<string> items = new List<string>();
        private string itemXmlElementName = XmlElementNames.String;

        /// <summary>
        /// Initializes a new instance of the <see cref="StringList"/> class.
        /// </summary>
        public StringList()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="StringList"/> class.
        /// </summary>
        /// <param name="strings">The strings.</param>
        public StringList(IEnumerable<string> strings)
        {
            this.AddRange(strings);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="StringList"/> class.
        /// </summary>
        /// <param name="itemXmlElementName">Name of the item XML element.</param>
        internal StringList(string itemXmlElementName)
        {
            this.itemXmlElementName = itemXmlElementName;
        }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            if (reader.LocalName == this.itemXmlElementName)
            {
                this.Add(reader.ReadValue());

                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Loads from json collection.
        /// </summary>
        /// <param name="jsonCollection">The json collection.</param>
        /// <param name="service">The service.</param>
        void IJsonCollectionDeserializer.CreateFromJsonCollection(object[] jsonCollection, ExchangeService service)
        {
            foreach (object element in jsonCollection)
            {
                this.Add(element as string);
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
        /// Writes elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            foreach (string item in this)
            {
                writer.WriteStartElement(XmlNamespace.Types, this.itemXmlElementName);
                writer.WriteValue(item, this.itemXmlElementName);
                writer.WriteEndElement();
            }
        }

        /// <summary>
        /// Serializes the property to a Json value.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <returns>
        /// A Json value (either a JsonObject, an array of Json values, or a Json primitive)
        /// </returns>
        internal override object InternalToJson(ExchangeService service)
        {
            return new List<string>(this).ToArray();
        }

        /// <summary>
        /// Adds a string to the list.
        /// </summary>
        /// <param name="s">The string to add.</param>
        public void Add(string s)
        {
            this.items.Add(s);
            this.Changed();
        }

        /// <summary>
        /// Adds multiple strings to the list.
        /// </summary>
        /// <param name="strings">The strings to add.</param>
        public void AddRange(IEnumerable<string> strings)
        {
            bool changed = false;

            foreach (string s in strings)
            {
                if (!this.Contains(s))
                {
                    this.items.Add(s);
                    changed = true;
                }
            }

            if (changed)
            {
                this.Changed();
            }
        }

        /// <summary>
        /// Determines whether the list contains a specific string.
        /// </summary>
        /// <param name="s">The string to check the presence of.</param>
        /// <returns>True if s is present in the list, false otherwise.</returns>
        public bool Contains(string s)
        {
            return this.items.Contains(s);
        }

        /// <summary>
        /// Removes a string from the list.
        /// </summary>
        /// <param name="s">The string to remove.</param>
        /// <returns>True is s was removed, false otherwise.</returns>
        public bool Remove(string s)
        {
            bool result = this.items.Remove(s);

            if (result)
            {
                this.Changed();
            }

            return result;
        }

        /// <summary>
        /// Removes the string at the specified position from the list.
        /// </summary>
        /// <param name="index">The index of the string to remove.</param>
        public void RemoveAt(int index)
        {
            if (index < 0 || index >= this.Count)
            {
                throw new ArgumentOutOfRangeException("index", Strings.IndexIsOutOfRange);
            }

            this.items.RemoveAt(index);

            this.Changed();
        }

        /// <summary>
        /// Clears the list.
        /// </summary>
        public void Clear()
        {
            this.items.Clear();
            this.Changed();
        }

        /// <summary>
        /// Generates a string representation of all the items in the list.
        /// </summary>
        /// <returns>A comma-separated list of the strings present in the list.</returns>
        public override string ToString()
        {
            return string.Join(",", this.items.ToArray());
        }

        /// <summary>
        /// Gets the number of strings in the list.
        /// </summary>
        public int Count
        {
            get { return this.items.Count; }
        }

        /// <summary>
        /// Gets or sets the string at the specified index.
        /// </summary>
        /// <param name="index">The index of the string to get or set.</param>
        /// <returns>The string at the specified index.</returns>
        public string this[int index]
        {
            get
            {
                if (index < 0 || index >= this.Count)
                {
                    throw new ArgumentOutOfRangeException("index", Strings.IndexIsOutOfRange);
                }

                return this.items[index];
            }

            set
            {
                if (index < 0 || index >= this.Count)
                {
                    throw new ArgumentOutOfRangeException("index", Strings.IndexIsOutOfRange);
                }

                if (this.items[index] != value)
                {
                    this.items[index] = value;
                    this.Changed();
                }
            }
        }

        #region IEnumerable<string> Members

        /// <summary>
        /// Gets an enumerator that iterates through the elements of the collection.
        /// </summary>
        /// <returns>An IEnumerator for the collection.</returns>
        public IEnumerator<string> GetEnumerator()
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

        /// <summary>
        /// Determines whether the specified <see cref="T:System.Object"/> is equal to the current <see cref="T:System.Object"/>.
        /// </summary>
        /// <param name="obj">The <see cref="T:System.Object"/> to compare with the current <see cref="T:System.Object"/>.</param>
        /// <returns>
        /// true if the specified <see cref="T:System.Object"/> is equal to the current <see cref="T:System.Object"/>; otherwise, false.
        /// </returns>
        /// <exception cref="T:System.NullReferenceException">The <paramref name="obj"/> parameter is null.</exception>
        public override bool Equals(object obj)
        {
            StringList other = obj as StringList;
            if (other != null)
            {
                return this.ToString().Equals(other.ToString());
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Serves as a hash function for a particular type.
        /// </summary>
        /// <returns>
        /// A hash code for the current <see cref="T:System.Object"/>.
        /// </returns>
        public override int GetHashCode()
        {
            return this.ToString().GetHashCode();
        }
    }
}
