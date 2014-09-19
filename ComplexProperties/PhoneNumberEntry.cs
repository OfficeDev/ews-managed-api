// ---------------------------------------------------------------------------
// <copyright file="PhoneNumberEntry.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the PhoneNumberEntry class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Text;

    /// <summary>
    /// Represents an entry of a PhoneNumberDictionary.
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public sealed class PhoneNumberEntry : DictionaryEntryProperty<PhoneNumberKey>
    {
        private string phoneNumber;

        /// <summary>
        /// Initializes a new instance of the <see cref="PhoneNumberEntry"/> class.
        /// </summary>
        internal PhoneNumberEntry()
            : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PhoneNumberEntry"/> class.
        /// </summary>
        /// <param name="key">The key.</param>
        /// <param name="phoneNumber">The phone number.</param>
        internal PhoneNumberEntry(PhoneNumberKey key, string phoneNumber)
            : base(key)
        {
            this.phoneNumber = phoneNumber;
        }

        /// <summary>
        /// Reads the text value from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadTextValueFromXml(EwsServiceXmlReader reader)
        {
            this.phoneNumber = reader.ReadValue();
        }

        /// <summary>
        /// Writes elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteValue(this.PhoneNumber, XmlElementNames.PhoneNumber);
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
            jsonProperty.Add(XmlElementNames.PhoneNumber, this.PhoneNumber);

            return jsonProperty;
        }

        /// <summary>
        /// Loads from json.
        /// </summary>
        /// <param name="jsonProperty">The json property.</param>
        /// <param name="service">The service.</param>
        internal override void LoadFromJson(JsonObject jsonProperty, ExchangeService service)
        {
            this.Key = jsonProperty.ReadEnumValue<PhoneNumberKey>(XmlAttributeNames.Key);
            this.PhoneNumber = jsonProperty.ReadAsString(XmlElementNames.PhoneNumber);
        }

        /// <summary>
        /// Gets or sets the phone number of the entry.
        /// </summary>
        public string PhoneNumber
        {
            get { return this.phoneNumber; }
            set { this.SetFieldValue<string>(ref this.phoneNumber, value); }
        }
    }
}
