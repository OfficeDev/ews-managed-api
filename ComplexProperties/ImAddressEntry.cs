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
// <summary>Defines the ImAddressEntry class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Text;

    /// <summary>
    /// Represents an entry of an ImAddressDictionary.
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public sealed class ImAddressEntry : DictionaryEntryProperty<ImAddressKey>
    {
        private string imAddress;

        /// <summary>
        /// Initializes a new instance of the <see cref="ImAddressEntry"/> class.
        /// </summary>
        internal ImAddressEntry()
            : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ImAddressEntry"/> class.
        /// </summary>
        /// <param name="key">The key.</param>
        /// <param name="imAddress">The im address.</param>
        internal ImAddressEntry(ImAddressKey key, string imAddress)
            : base(key)
        {
            this.imAddress = imAddress;
        }

        /// <summary>
        /// Gets or sets the Instant Messaging address of the entry.
        /// </summary>
        public string ImAddress
        {
            get { return this.imAddress; }
            set { this.SetFieldValue<string>(ref this.imAddress, value); }
        }

        /// <summary>
        /// Reads the text value from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadTextValueFromXml(EwsServiceXmlReader reader)
        {
            this.imAddress = reader.ReadValue();
        }

        /// <summary>
        /// Writes elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteValue(this.ImAddress, XmlElementNames.ImAddress);
        }

        internal override object InternalToJson(ExchangeService service)
        {
            JsonObject jsonProperty = new JsonObject();

            jsonProperty.Add(XmlAttributeNames.Key, this.Key);
            jsonProperty.Add(XmlElementNames.ImAddress, this.ImAddress);

            return jsonProperty;
        }

        /// <summary>
        /// Loads from json.
        /// </summary>
        /// <param name="jsonProperty">The json property.</param>
        /// <param name="service">The service.</param>
        internal override void LoadFromJson(JsonObject jsonProperty, ExchangeService service)
        {
            this.Key = jsonProperty.ReadEnumValue<ImAddressKey>(XmlAttributeNames.Key);
            this.ImAddress = jsonProperty.ReadAsString(XmlElementNames.ImAddress);
        }
    }
}
