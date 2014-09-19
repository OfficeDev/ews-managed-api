// ---------------------------------------------------------------------------
// <copyright file="ComplexPropertyCollection.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ComplexPropertyCollection class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Text;

    /// <summary>
    /// Represents a collection of properties that can be sent to and retrieved from EWS.
    /// </summary>
    /// <typeparam name="TComplexProperty">ComplexProperty type.</typeparam>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public abstract class ComplexPropertyCollection<TComplexProperty> : ComplexProperty, IEnumerable<TComplexProperty>, ICustomUpdateSerializer, IJsonCollectionDeserializer
        where TComplexProperty : ComplexProperty
    {
        private List<TComplexProperty> items = new List<TComplexProperty>();
        private List<TComplexProperty> addedItems = new List<TComplexProperty>();
        private List<TComplexProperty> modifiedItems = new List<TComplexProperty>();
        private List<TComplexProperty> removedItems = new List<TComplexProperty>();

        /// <summary>
        /// Creates the complex property.
        /// </summary>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <returns>Complex property instance.</returns>
        internal abstract TComplexProperty CreateComplexProperty(string xmlElementName);

        /// <summary>
        /// Creates the default complex property.
        /// </summary>
        /// <returns>Complex property instance.</returns>
        internal abstract TComplexProperty CreateDefaultComplexProperty();

        /// <summary>
        /// Gets the name of the collection item XML element.
        /// </summary>
        /// <param name="complexProperty">The complex property.</param>
        /// <returns>XML element name.</returns>
        internal abstract string GetCollectionItemXmlElementName(TComplexProperty complexProperty);

        /// <summary>
        /// Initializes a new instance of the <see cref="ComplexPropertyCollection&lt;TComplexProperty&gt;"/> class.
        /// </summary>
        internal ComplexPropertyCollection()
            : base()
        {
        }

        /// <summary>
        /// Item changed.
        /// </summary>
        /// <param name="complexProperty">The complex property.</param>
        internal void ItemChanged(ComplexProperty complexProperty)
        {
            TComplexProperty property = complexProperty as TComplexProperty;

            EwsUtilities.Assert(
                property != null,
                "ComplexPropertyCollection.ItemChanged",
                string.Format("ComplexPropertyCollection.ItemChanged: the type of the complexProperty argument ({0}) is not supported.", complexProperty.GetType().Name));

            if (!this.addedItems.Contains(property))
            {
                if (!this.modifiedItems.Contains(property))
                {
                    this.modifiedItems.Add(property);
                    this.Changed();
                }
            }
        }

        /// <summary>
        /// Loads from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <param name="localElementName">Name of the local element.</param>
        internal override void LoadFromXml(EwsServiceXmlReader reader, string localElementName)
        {
            this.LoadFromXml(
                reader,
                XmlNamespace.Types,
                localElementName);
        }

        /// <summary>
        /// Loads from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <param name="xmlNamespace">The XML namespace.</param>
        /// <param name="localElementName">Name of the local element.</param>
        internal override void LoadFromXml(
            EwsServiceXmlReader reader,
            XmlNamespace xmlNamespace,
            string localElementName)
        {
            reader.EnsureCurrentNodeIsStartElement(xmlNamespace, localElementName);

            if (!reader.IsEmptyElement)
            {
                do
                {
                    reader.Read();

                    if (reader.IsStartElement())
                    {
                        TComplexProperty complexProperty = this.CreateComplexProperty(reader.LocalName);

                        if (complexProperty != null)
                        {
                            complexProperty.LoadFromXml(reader, reader.LocalName);
                            this.InternalAdd(complexProperty, true);
                        }
                        else
                        {
                            reader.SkipCurrentElement();
                        }
                    }
                }
                while (!reader.IsEndElement(xmlNamespace, localElementName));
            }
        }

        /// <summary>
        /// Loads from XML to update itself.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <param name="xmlNamespace">The XML namespace.</param>
        /// <param name="xmlElementName">Name of the XML element.</param>
        internal override void UpdateFromXml(
            EwsServiceXmlReader reader,
            XmlNamespace xmlNamespace,
            string xmlElementName)
        {
            reader.EnsureCurrentNodeIsStartElement(xmlNamespace, xmlElementName);

            if (!reader.IsEmptyElement)
            {
                int index = 0;
                do
                {
                    reader.Read();

                    if (reader.IsStartElement())
                    {
                        TComplexProperty complexProperty = this.CreateComplexProperty(reader.LocalName);
                        TComplexProperty actualComplexProperty = this[index++];

                        if (complexProperty == null || !complexProperty.GetType().IsInstanceOfType(actualComplexProperty))
                        {
                            throw new ServiceLocalException(Strings.PropertyTypeIncompatibleWhenUpdatingCollection);
                        }

                        actualComplexProperty.UpdateFromXml(reader, xmlNamespace, reader.LocalName);
                    }
                }
                while (!reader.IsEndElement(xmlNamespace, xmlElementName));
            }
        }

        /// <summary>
        /// Loads from json.
        /// </summary>
        /// <param name="jsonCollection">The json collection.</param>
        /// <param name="service">The service.</param>
        void IJsonCollectionDeserializer.CreateFromJsonCollection(object[] jsonCollection, ExchangeService service)
        {
            foreach (object jsonObject in jsonCollection)
            {
                JsonObject jsonProperty = jsonObject as JsonObject;

                if (jsonProperty != null)
                {
                    TComplexProperty complexProperty = null;

                    // If type property is present, use it. Otherwise create default property instance.
                    // Note: polymorphic collections (such as Attachments) need a type property so
                    // the CreateDefaultComplexProperty call will fail.
                    if (jsonProperty.HasTypeProperty())
                    {
                        complexProperty = this.CreateComplexProperty(jsonProperty.ReadTypeString());
                    }
                    else
                    {
                        complexProperty = this.CreateDefaultComplexProperty();
                    }

                    if (complexProperty != null)
                    {
                        complexProperty.LoadFromJson(jsonProperty, service);
                        this.InternalAdd(complexProperty, true);
                    }
                }
            }
        }

        /// <summary>
        /// Loads from json to update existing property.
        /// </summary>
        /// <param name="jsonCollection">The json collection.</param>
        /// <param name="service">The service.</param>
        void IJsonCollectionDeserializer.UpdateFromJsonCollection(object[] jsonCollection, ExchangeService service)
        {
            if (this.Count != jsonCollection.Length)
            {
                throw new ServiceLocalException(Strings.PropertyCollectionSizeMismatch);
            }

            int index = 0;
            foreach (object jsonObject in jsonCollection)
            {
                JsonObject jsonProperty = jsonObject as JsonObject;

                if (jsonProperty != null)
                {
                    TComplexProperty expectedComplexProperty = null;

                    if (jsonProperty.HasTypeProperty())
                    {
                        expectedComplexProperty = this.CreateComplexProperty(jsonProperty.ReadTypeString());
                    }
                    else
                    {
                        expectedComplexProperty = this.CreateDefaultComplexProperty();
                    }

                    TComplexProperty actualComplexProperty = this[index++];

                    if (expectedComplexProperty == null || !expectedComplexProperty.GetType().IsInstanceOfType(actualComplexProperty))
                    {
                        throw new ServiceLocalException(Strings.PropertyTypeIncompatibleWhenUpdatingCollection);
                    }

                    actualComplexProperty.LoadFromJson(jsonProperty, service);
                }
                else
                {
                    throw new ServiceLocalException();
                }
            }
        }

        /// <summary>
        /// Writes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="xmlNamespace">The XML namespace.</param>
        /// <param name="xmlElementName">Name of the XML element.</param>
        internal override void WriteToXml(
            EwsServiceXmlWriter writer,
            XmlNamespace xmlNamespace,
            string xmlElementName)
        {
            if (this.ShouldWriteToRequest())
            {
                base.WriteToXml(
                    writer,
                    xmlNamespace,
                    xmlElementName);
            }
        }

        /// <summary>
        /// Serializes the property to a Json value.
        /// </summary>
        /// <param name="service"></param>
        /// <returns>
        /// A Json value (either a JsonObject, an array of Json values, or a Json primitive)
        /// </returns>
        internal override object InternalToJson(ExchangeService service)
        {
            List<object> jsonPropertyCollection = new List<object>();

            foreach (TComplexProperty complexProperty in this)
            {
                jsonPropertyCollection.Add(complexProperty.InternalToJson(service));
            }

            return jsonPropertyCollection.ToArray();
        }

        /// <summary>
        /// Determine whether we should write collection to XML or not.
        /// </summary>
        /// <returns>True if collection contains at least one element.</returns>
        internal virtual bool ShouldWriteToRequest()
        {
            // Only write collection if it has at least one element.
            return this.Count > 0;
        }

        /// <summary>
        /// Writes elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            foreach (TComplexProperty complexProperty in this)
            {
                complexProperty.WriteToXml(writer, this.GetCollectionItemXmlElementName(complexProperty));
            }
        }

        /// <summary>
        /// Clears the change log.
        /// </summary>
        internal override void ClearChangeLog()
        {
            this.removedItems.Clear();
            this.addedItems.Clear();
            this.modifiedItems.Clear();
        }

        /// <summary>
        /// Removes from change log.
        /// </summary>
        /// <param name="complexProperty">The complex property.</param>
        internal void RemoveFromChangeLog(TComplexProperty complexProperty)
        {
            this.removedItems.Remove(complexProperty);
            this.modifiedItems.Remove(complexProperty);
            this.addedItems.Remove(complexProperty);
        }

        /// <summary>
        /// Gets the items.
        /// </summary>
        /// <value>The items.</value>
        internal List<TComplexProperty> Items
        {
            get { return this.items; }
        }

        /// <summary>
        /// Gets the added items.
        /// </summary>
        /// <value>The added items.</value>
        internal List<TComplexProperty> AddedItems
        {
            get { return this.addedItems; }
        }

        /// <summary>
        /// Gets the modified items.
        /// </summary>
        /// <value>The modified items.</value>
        internal List<TComplexProperty> ModifiedItems
        {
            get { return this.modifiedItems; }
        }

        /// <summary>
        /// Gets the removed items.
        /// </summary>
        /// <value>The removed items.</value>
        internal List<TComplexProperty> RemovedItems
        {
            get { return this.removedItems; }
        }

        /// <summary>
        /// Add complex property.
        /// </summary>
        /// <param name="complexProperty">The complex property.</param>
        internal void InternalAdd(TComplexProperty complexProperty)
        {
            this.InternalAdd(complexProperty, false);
        }

        /// <summary>
        /// Add complex property.
        /// </summary>
        /// <param name="complexProperty">The complex property.</param>
        /// <param name="loading">If true, collection is being loaded.</param>
        private void InternalAdd(TComplexProperty complexProperty, bool loading)
        {
            EwsUtilities.Assert(
                complexProperty != null,
                "ComplexPropertyCollection.InternalAdd",
                "complexProperty is null");

            if (!this.items.Contains(complexProperty))
            {
                this.items.Add(complexProperty);
                if (!loading)
                {
                    this.removedItems.Remove(complexProperty);
                    this.addedItems.Add(complexProperty);
                }
                complexProperty.OnChange += this.ItemChanged;
                this.Changed();
            }
        }

        /// <summary>
        /// Clear collection.
        /// </summary>
        internal void InternalClear()
        {
            while (this.Count > 0)
            {
                this.InternalRemoveAt(0);
            }
        }

        /// <summary>
        /// Remote entry at index.
        /// </summary>
        /// <param name="index">The index.</param>
        internal void InternalRemoveAt(int index)
        {
            EwsUtilities.Assert(
              index >= 0 && index < this.Count,
              "ComplexPropertyCollection.InternalRemoveAt",
              "index is out of range.");

            this.InternalRemove(this.items[index]);
        }

        /// <summary>
        /// Remove specified complex property.
        /// </summary>
        /// <param name="complexProperty">The complex property.</param>
        /// <returns>True if the complex property was successfully removed from the collection, false otherwise.</returns>
        internal bool InternalRemove(TComplexProperty complexProperty)
        {
            EwsUtilities.Assert(
                complexProperty != null,
                "ComplexPropertyCollection.InternalRemove",
                "complexProperty is null");

            if (this.items.Remove(complexProperty))
            {
                complexProperty.OnChange -= this.ItemChanged;

                if (!this.addedItems.Contains(complexProperty))
                {
                    this.removedItems.Add(complexProperty);
                }
                else
                {
                    this.addedItems.Remove(complexProperty);
                }
                this.modifiedItems.Remove(complexProperty);
                this.Changed();
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Determines whether a specific property is in the collection.
        /// </summary>
        /// <param name="complexProperty">The property to locate in the collection.</param>
        /// <returns>True if the property was found in the collection, false otherwise.</returns>
        public bool Contains(TComplexProperty complexProperty)
        {
            return this.items.Contains(complexProperty);
        }

        /// <summary>
        /// Searches for a specific property and return its zero-based index within the collection.
        /// </summary>
        /// <param name="complexProperty">The property to locate in the collection.</param>
        /// <returns>The zero-based index of the property within the collection.</returns>
        public int IndexOf(TComplexProperty complexProperty)
        {
            return this.items.IndexOf(complexProperty);
        }

        /// <summary>
        /// Gets the total number of properties in the collection.
        /// </summary>
        public int Count
        {
            get { return this.items.Count; }
        }

        /// <summary>
        /// Gets the property at the specified index.
        /// </summary>
        /// <param name="index">The zero-based index of the property to get.</param>
        /// <returns>The property at the specified index.</returns>
        public TComplexProperty this[int index]
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

        #region IEnumerable<TComplexProperty> Members

        /// <summary>
        /// Gets an enumerator that iterates through the elements of the collection.
        /// </summary>
        /// <returns>An IEnumerator for the collection.</returns>
        public IEnumerator<TComplexProperty> GetEnumerator()
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

        #region ICustomXmlUpdateSerializer Members

        /// <summary>
        /// Writes the update to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="ewsObject">The ews object.</param>
        /// <param name="propertyDefinition">Property definition.</param>
        /// <returns>True if property generated serialization.</returns>
        bool ICustomUpdateSerializer.WriteSetUpdateToXml(
            EwsServiceXmlWriter writer,
            ServiceObject ewsObject,
            PropertyDefinition propertyDefinition)
        {
            // If the collection is empty, delete the property.
            if (this.Count == 0)
            {
                writer.WriteStartElement(XmlNamespace.Types, ewsObject.GetDeleteFieldXmlElementName());
                propertyDefinition.WriteToXml(writer);
                writer.WriteEndElement();
                return true;
            }

            // Otherwise, use the default XML serializer.
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Writes the deletion update to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="ewsObject">The ews object.</param>
        /// <returns>True if property generated serialization.</returns>
        bool ICustomUpdateSerializer.WriteDeleteUpdateToXml(EwsServiceXmlWriter writer, ServiceObject ewsObject)
        {
            // Use the default XML serializer.
            return false;
        }

        /// <summary>
        /// Writes the update to Json.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="ewsObject">The ews object.</param>
        /// <param name="propertyDefinition">Property definition.</param>
        /// <param name="updates">The updates.</param>
        /// <returns>
        /// True if property generated serialization.
        /// </returns>
        bool ICustomUpdateSerializer.WriteSetUpdateToJson(ExchangeService service, ServiceObject ewsObject, PropertyDefinition propertyDefinition, List<JsonObject> updates)
        {
            // If the collection is empty, delete the property.
            if (this.Count == 0)
            {
                JsonObject jsonUpdate = new JsonObject();

                jsonUpdate.AddTypeParameter(ewsObject.GetDeleteFieldXmlElementName());
                jsonUpdate.Add(JsonNames.Path, (propertyDefinition as IJsonSerializable).ToJson(service));
                return true;
            }

            // Otherwise, use the default Json serializer.
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Writes the deletion update to Json.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="ewsObject">The ews object.</param>
        /// <param name="updates">The updates.</param>
        /// <returns>
        /// True if property generated serialization.
        /// </returns>
        bool ICustomUpdateSerializer.WriteDeleteUpdateToJson(ExchangeService service, ServiceObject ewsObject, List<JsonObject> updates)
        {
            // Use the default Json serializer.
            return false;
        }

        #endregion
    }
}
