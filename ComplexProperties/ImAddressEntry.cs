// ---------------------------------------------------------------------------
// <copyright file="ImAddressEntry.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

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
