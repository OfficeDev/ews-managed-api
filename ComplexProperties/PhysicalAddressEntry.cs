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
// <summary>Implements the representation of a physical address.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents an entry of an PhysicalAddressDictionary.
    /// </summary>
    public sealed class PhysicalAddressEntry : DictionaryEntryProperty<PhysicalAddressKey>
    {
        #region Fields

        private SimplePropertyBag<string> propertyBag;

        #endregion

        #region Constructors

        /// <summary>
        /// Initializes a new instance of PhysicalAddressEntry
        /// </summary>
        public PhysicalAddressEntry()
            : base()
        {
            this.propertyBag = new SimplePropertyBag<string>();
            this.propertyBag.OnChange += this.PropertyBagChanged;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Gets or sets the street.
        /// </summary>
        public string Street
        {
            get { return (string)this.propertyBag[PhysicalAddressSchema.Street]; }
            set { this.propertyBag[PhysicalAddressSchema.Street] = value; }
        }

        /// <summary>
        /// Gets or sets the city.
        /// </summary>
        public string City
        {
            get { return (string)this.propertyBag[PhysicalAddressSchema.City]; }
            set { this.propertyBag[PhysicalAddressSchema.City] = value; }
        }

        /// <summary>
        /// Gets or sets the state.
        /// </summary>
        public string State
        {
            get { return (string)this.propertyBag[PhysicalAddressSchema.State]; }
            set { this.propertyBag[PhysicalAddressSchema.State] = value; }
        }

        /// <summary>
        /// Gets or sets the country or region.
        /// </summary>
        public string CountryOrRegion
        {
            get { return (string)this.propertyBag[PhysicalAddressSchema.CountryOrRegion]; }
            set { this.propertyBag[PhysicalAddressSchema.CountryOrRegion] = value; }
        }

        /// <summary>
        /// Gets or sets the postal code.
        /// </summary>
        public string PostalCode
        {
            get { return (string)this.propertyBag[PhysicalAddressSchema.PostalCode]; }
            set { this.propertyBag[PhysicalAddressSchema.PostalCode] = value; }
        }

        #endregion

        #region Internal methods

        /// <summary>
        /// Clears the change log.
        /// </summary>
        internal override void ClearChangeLog()
        {
            this.propertyBag.ClearChangeLog();
        }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            if (PhysicalAddressSchema.XmlElementNames.Contains(reader.LocalName))
            {
                this.propertyBag[reader.LocalName] = reader.ReadElementValue();

                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Writes elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            foreach (string xmlElementName in PhysicalAddressSchema.XmlElementNames)
            {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    xmlElementName,
                    this.propertyBag[xmlElementName]);
            }
        }

        /// <summary>
        /// Writes the update to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="ewsObject">The ews object.</param>
        /// <param name="ownerDictionaryXmlElementName">Name of the owner dictionary XML element.</param>
        /// <returns>True if update XML was written.</returns>
        internal override bool WriteSetUpdateToXml(
            EwsServiceXmlWriter writer,
            ServiceObject ewsObject,
            string ownerDictionaryXmlElementName)
        {
            List<string> fieldsToSet = new List<string>();

            foreach (string xmlElementName in this.propertyBag.AddedItems)
            {
                fieldsToSet.Add(xmlElementName);
            }

            foreach (string xmlElementName in this.propertyBag.ModifiedItems)
            {
                fieldsToSet.Add(xmlElementName);
            }

            foreach (string xmlElementName in fieldsToSet)
            {
                writer.WriteStartElement(XmlNamespace.Types, ewsObject.GetSetFieldXmlElementName());

                writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.IndexedFieldURI);
                writer.WriteAttributeValue(XmlAttributeNames.FieldURI, GetFieldUri(xmlElementName));
                writer.WriteAttributeValue(XmlAttributeNames.FieldIndex, this.Key.ToString());
                writer.WriteEndElement(); // IndexedFieldURI

                writer.WriteStartElement(XmlNamespace.Types, ewsObject.GetXmlElementName());
                writer.WriteStartElement(XmlNamespace.Types, ownerDictionaryXmlElementName);
                writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.Entry);
                this.WriteAttributesToXml(writer);
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    xmlElementName,
                    this.propertyBag[xmlElementName]);
                writer.WriteEndElement(); // Entry
                writer.WriteEndElement(); // ownerDictionaryXmlElementName
                writer.WriteEndElement(); // ewsObject.GetXmlElementName()

                writer.WriteEndElement(); // ewsObject.GetSetFieldXmlElementName()
            }

            foreach (string xmlElementName in this.propertyBag.RemovedItems)
            {
                this.InternalWriteDeleteFieldToXml(
                    writer,
                    ewsObject,
                    xmlElementName);
            }

            return true;
        }

        /// <summary>
        /// Writes the set update to json.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="ewsObject">The ews object.</param>
        /// <param name="propertyDefinition">The property definition.</param>
        /// <param name="updates">The updates.</param>
        /// <returns></returns>
        internal override bool WriteSetUpdateToJson(ExchangeService service, ServiceObject ewsObject, PropertyDefinition propertyDefinition, List<JsonObject> updates)
        {
            List<string> propertiesToSet = new List<string>();

            propertiesToSet.AddRange(this.propertyBag.AddedItems);
            propertiesToSet.AddRange(this.propertyBag.ModifiedItems);

            foreach (string propertyName in propertiesToSet)
            {
                JsonObject jsonUpdate = new JsonObject();

                jsonUpdate.AddTypeParameter(ewsObject.GetSetFieldXmlElementName());

                JsonObject jsonUri = new JsonObject();

                jsonUri.AddTypeParameter(JsonNames.PathToIndexedFieldType);
                jsonUri.Add(XmlAttributeNames.FieldURI, GetFieldUri(propertyName));
                jsonUri.Add(XmlAttributeNames.FieldIndex, this.Key.ToString());

                jsonUpdate.Add(JsonNames.Path, jsonUri);

                JsonObject jsonProperty = new JsonObject();
                jsonProperty.Add(XmlAttributeNames.Key, this.Key);
                jsonProperty.Add(propertyName, this.propertyBag[propertyName]);

                JsonObject jsonServiceObject = new JsonObject();
                jsonServiceObject.AddTypeParameter(ewsObject.GetXmlElementName());
                jsonServiceObject.Add(propertyDefinition.XmlElementName, new JsonObject[] { jsonProperty });

                jsonUpdate.Add(PropertyBag.GetPropertyUpdateItemName(ewsObject), jsonServiceObject);

                updates.Add(jsonUpdate);
            }

            foreach (string propertyName in this.propertyBag.RemovedItems)
            {
                this.InternalWriteDeleteUpdateToJson(
                                    ewsObject,
                                    propertyName,
                                    updates);
            }

            return true;
        }

        /// <summary>
        /// Writes the delete update to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="ewsObject">The ews object.</param>
        /// <returns>True if update XML was written.</returns>
        internal override bool WriteDeleteUpdateToXml(EwsServiceXmlWriter writer, ServiceObject ewsObject)
        {
            foreach (string xmlElementName in PhysicalAddressSchema.XmlElementNames)
            {
                this.InternalWriteDeleteFieldToXml(
                    writer,
                    ewsObject,
                    xmlElementName);
            }

            return true;
        }

        /// <summary>
        /// Writes field deletion to json.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="ewsObject">The ews object.</param>
        /// <param name="updates">The updates.</param>
        /// <returns></returns>
        internal override bool WriteDeleteUpdateToJson(ExchangeService service, ServiceObject ewsObject, List<JsonObject> updates)
        {
            foreach (string propertyName in PhysicalAddressSchema.XmlElementNames)
            {
                this.InternalWriteDeleteUpdateToJson(
                    ewsObject,
                    propertyName,
                    updates);
            }

            return true;
        }

        /// <summary>
        /// Internal the write delete update to json.
        /// </summary>
        /// <param name="ewsObject">The ews object.</param>
        /// <param name="propertyName">Name of the property.</param>
        /// <param name="updates">The updates.</param>
        internal void InternalWriteDeleteUpdateToJson(ServiceObject ewsObject, string propertyName, List<JsonObject> updates)
        {
            JsonObject jsonUpdate = new JsonObject();

            jsonUpdate.AddTypeParameter(ewsObject.GetDeleteFieldXmlElementName());

            JsonObject jsonUri = new JsonObject();

            jsonUri.AddTypeParameter(JsonNames.PathToIndexedFieldType);
            jsonUri.Add(XmlAttributeNames.FieldURI, GetFieldUri(propertyName));
            jsonUri.Add(XmlAttributeNames.FieldIndex, this.Key.ToString());

            jsonUpdate.Add(JsonNames.Path, jsonUri);

            updates.Add(jsonUpdate);
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
            JsonObject jsonProperty = new JsonObject();

            jsonProperty.Add(XmlAttributeNames.Key, this.Key);
            jsonProperty.Add(XmlElementNames.Street, this.Street);
            jsonProperty.Add(XmlElementNames.City, this.City);
            jsonProperty.Add(XmlElementNames.State, this.State);
            jsonProperty.Add(XmlElementNames.CountryOrRegion, this.CountryOrRegion);
            jsonProperty.Add(XmlElementNames.PostalCode, this.PostalCode);

            return jsonProperty;
        }

        /// <summary>
        /// Loads from json.
        /// </summary>
        /// <param name="jsonProperty">The json property.</param>
        /// <param name="service">The service.</param>
        internal override void LoadFromJson(JsonObject jsonProperty, ExchangeService service)
        {
            this.Key = jsonProperty.ReadEnumValue<PhysicalAddressKey>(XmlAttributeNames.Key);
            this.Street = jsonProperty.ReadAsString(XmlElementNames.Street);
            this.City = jsonProperty.ReadAsString(XmlElementNames.City);
            this.State = jsonProperty.ReadAsString(XmlElementNames.State);
            this.Street = jsonProperty.ReadAsString(XmlElementNames.Street);
            this.CountryOrRegion = jsonProperty.ReadAsString(XmlElementNames.CountryOrRegion);
            this.PostalCode = jsonProperty.ReadAsString(XmlElementNames.PostalCode);
        }

        #endregion

        #region Private methods

        /// <summary>
        /// Gets the field URI.
        /// </summary>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <returns>Field URI.</returns>
        private static string GetFieldUri(string xmlElementName)
        {
            return "contacts:PhysicalAddress:" + xmlElementName;
        }

        /// <summary>
        /// Property bag was changed.
        /// </summary>
        private void PropertyBagChanged()
        {
            this.Changed();
        }

        /// <summary>
        /// Write field deletion to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="ewsObject">The ews object.</param>
        /// <param name="fieldXmlElementName">Name of the field XML element.</param>
        private void InternalWriteDeleteFieldToXml(
            EwsServiceXmlWriter writer,
            ServiceObject ewsObject,
            string fieldXmlElementName)
        {
            writer.WriteStartElement(XmlNamespace.Types, ewsObject.GetDeleteFieldXmlElementName());
            writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.IndexedFieldURI);
            writer.WriteAttributeValue(XmlAttributeNames.FieldURI, GetFieldUri(fieldXmlElementName));
            writer.WriteAttributeValue(XmlAttributeNames.FieldIndex, this.Key.ToString());
            writer.WriteEndElement(); // IndexedFieldURI
            writer.WriteEndElement(); // ewsObject.GetDeleteFieldXmlElementName()
        }

        #endregion

        #region Classes

        /// <summary>
        /// Schema definition for PhysicalAddress
        /// </summary>
        private static class PhysicalAddressSchema
        {
            public const string Street = "Street";
            public const string City = "City";
            public const string State = "State";
            public const string CountryOrRegion = "CountryOrRegion";
            public const string PostalCode = "PostalCode";

            /// <summary>
            /// List of XML element names.
            /// </summary>
            private static LazyMember<List<string>> xmlElementNames = new LazyMember<List<string>>(
                delegate()
                {
                    List<string> result = new List<string>();
                    result.Add(Street);
                    result.Add(City);
                    result.Add(State);
                    result.Add(CountryOrRegion);
                    result.Add(PostalCode);
                    return result;
                });

            /// <summary>
            /// Gets the XML element names.
            /// </summary>
            /// <value>The XML element names.</value>
            public static List<string> XmlElementNames
            {
                get { return xmlElementNames.Member; }
            }
        }

        #endregion
    }
}
