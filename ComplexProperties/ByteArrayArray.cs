// ---------------------------------------------------------------------------
// <copyright file="ByteArrayArray.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ByteArrayArray class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Text;

    /// <summary>
    /// Represents an array of byte arrays
    /// </summary>
    public sealed class ByteArrayArray : ComplexProperty,  IJsonCollectionDeserializer
    {
        private const string ItemXmlElementName = "Base64Binary";
        private List<byte[]> content = new List<byte[]>();

        #region Properties

        /// <summary>
        /// Gets the content of the arrray of byte arrays
        /// </summary>
        public byte[][] Content
        {
            get { return this.content.ToArray(); }
        }

        #endregion

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            if (reader.LocalName == ByteArrayArray.ItemXmlElementName)
            {
                this.content.Add(reader.ReadBase64ElementValue());
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
                this.content.Add(Convert.FromBase64String(element as string));
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
        /// Writes the elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            foreach (byte[] item in this.content)
            {
                writer.WriteStartElement(XmlNamespace.Types, ByteArrayArray.ItemXmlElementName);
                writer.WriteBase64ElementValue(item);
                writer.WriteEndElement();
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
            List<string> base64Strings = new List<string>(this.content.Count);
            foreach (byte[] item in this.content)
            {
                base64Strings.Add(Convert.ToBase64String(item));
            }
            return base64Strings.ToArray();
        }
    }
}
