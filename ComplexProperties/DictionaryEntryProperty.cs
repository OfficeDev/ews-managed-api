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
    using System.ComponentModel;
    using System.Text;

    /// <summary>
    /// Represents an entry of a DictionaryProperty object.
    /// </summary>
    /// <remarks>
    /// All descendants of DictionaryEntryProperty must implement a parameterless
    /// constructor. That constructor does not have to be public.
    /// </remarks>
    /// <typeparam name="TKey">The type of the key used by this dictionary.</typeparam>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public abstract class DictionaryEntryProperty<TKey> : ComplexProperty
    {
        private TKey key;

        /// <summary>
        /// Initializes a new instance of the <see cref="DictionaryEntryProperty&lt;TKey&gt;"/> class.
        /// </summary>
        internal DictionaryEntryProperty()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DictionaryEntryProperty&lt;TKey&gt;"/> class.
        /// </summary>
        /// <param name="key">The key.</param>
        internal DictionaryEntryProperty(TKey key)
            : base()
        {
            this.key = key;
        }

        /// <summary>
        /// Gets or sets the key.
        /// </summary>
        /// <value>The key.</value>
        internal TKey Key
        {
            get { return this.key; }
            set { this.key = value; }
        }

        /// <summary>
        /// Reads the attributes from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadAttributesFromXml(EwsServiceXmlReader reader)
        {
            this.key = reader.ReadAttributeValue<TKey>(XmlAttributeNames.Key);
        }

        /// <summary>
        /// Writes the attributes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteAttributeValue(XmlAttributeNames.Key, this.Key);
        }

        /// <summary>
        /// Writes the set update to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="ewsObject">The ews object.</param>
        /// <param name="ownerDictionaryXmlElementName">Name of the owner dictionary XML element.</param>
        /// <returns>True if update XML was written.</returns>
        internal virtual bool WriteSetUpdateToXml(
            EwsServiceXmlWriter writer,
            ServiceObject ewsObject,
            string ownerDictionaryXmlElementName)
        {
            return false;
        }

        /// <summary>
        /// Writes the set update to json.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="ewsObject">The ews object.</param>
        /// <param name="propertyDefinition">The property definition.</param>
        /// <param name="updates">The updates.</param>
        /// <returns></returns>
        internal virtual bool WriteSetUpdateToJson(
             ExchangeService service,
             ServiceObject ewsObject,
             PropertyDefinition propertyDefinition,
             List<JsonObject> updates)
        {
            return false;
        }

        /// <summary>
        /// Writes the delete update to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="ewsObject">The ews object.</param>
        /// <returns>True if update XML was written.</returns>
        internal virtual bool WriteDeleteUpdateToXml(EwsServiceXmlWriter writer, ServiceObject ewsObject)
        {
            return false;
        }

        /// <summary>
        /// Writes the delete update to json.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="ewsObject">The ews object.</param>
        /// <param name="updates">The updates.</param>
        /// <returns></returns>
        internal virtual bool WriteDeleteUpdateToJson(ExchangeService service, ServiceObject ewsObject, List<JsonObject> updates)
        {
            return false;
        }
    }
}