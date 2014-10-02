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
// <summary>Defines the EmailAddressEntry class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System.ComponentModel;

    /// <summary>
    /// Represents an entry of an EmailAddressDictionary.
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public sealed class EmailAddressEntry : DictionaryEntryProperty<EmailAddressKey>
    {
        /// <summary>
        /// The email address.
        /// </summary>
        private EmailAddress emailAddress;

        /// <summary>
        /// Initializes a new instance of the <see cref="EmailAddressEntry"/> class.
        /// </summary>
        internal EmailAddressEntry()
            : base()
        {
            this.emailAddress = new EmailAddress();
            this.emailAddress.OnChange += this.EmailAddressChanged;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="EmailAddressEntry"/> class.
        /// </summary>
        /// <param name="key">The key.</param>
        /// <param name="emailAddress">The email address.</param>
        internal EmailAddressEntry(EmailAddressKey key, EmailAddress emailAddress)
            : base(key)
        {
            this.emailAddress = emailAddress;

            if (this.emailAddress != null)
            {
                this.emailAddress.OnChange += this.EmailAddressChanged;
            }
        }

        /// <summary>
        /// Reads the attributes from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadAttributesFromXml(EwsServiceXmlReader reader)
        {
            base.ReadAttributesFromXml(reader);

            this.EmailAddress.Name = reader.ReadAttributeValue<string>(XmlAttributeNames.Name);
            this.EmailAddress.RoutingType = reader.ReadAttributeValue<string>(XmlAttributeNames.RoutingType);

            string mailboxTypeString = reader.ReadAttributeValue(XmlAttributeNames.MailboxType);
            if (!string.IsNullOrEmpty(mailboxTypeString))
            {
                this.EmailAddress.MailboxType = EwsUtilities.Parse<MailboxType>(mailboxTypeString);
            }
            else
            {
                this.EmailAddress.MailboxType = null;
            }
        }

        /// <summary>
        /// Reads the text value from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadTextValueFromXml(EwsServiceXmlReader reader)
        {
            this.EmailAddress.Address = reader.ReadValue();
        }

        /// <summary>
        /// Loads from json.
        /// </summary>
        /// <param name="jsonProperty">The json property.</param>
        /// <param name="service">The service.</param>
        internal override void LoadFromJson(JsonObject jsonProperty, ExchangeService service)
        {
            foreach (string key in jsonProperty.Keys)
            {
                switch (key)
                {
                    case XmlAttributeNames.Key:
                        this.Key = jsonProperty.ReadEnumValue<EmailAddressKey>(key);
                        break;
                    case XmlAttributeNames.Name:
                        this.EmailAddress.Name = jsonProperty.ReadAsString(key);
                        break;
                    case XmlAttributeNames.RoutingType:
                        this.EmailAddress.RoutingType = jsonProperty.ReadAsString(key);
                        break;
                    case XmlAttributeNames.MailboxType:
                        this.EmailAddress.MailboxType = jsonProperty.ReadEnumValue<MailboxType>(key);
                        break;
                    case XmlElementNames.EmailAddress:
                        this.EmailAddress.Address = jsonProperty.ReadAsString(key);
                        break;
                }
            }
        }

        /// <summary>
        /// Writes the attributes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
        {
            base.WriteAttributesToXml(writer);

            if (writer.Service.RequestedServerVersion > ExchangeVersion.Exchange2007_SP1)
            {
                writer.WriteAttributeValue(XmlAttributeNames.Name, this.EmailAddress.Name);
                writer.WriteAttributeValue(XmlAttributeNames.RoutingType, this.EmailAddress.RoutingType);
                if (this.EmailAddress.MailboxType != MailboxType.Unknown)
                {
                    writer.WriteAttributeValue(XmlAttributeNames.MailboxType, this.EmailAddress.MailboxType);
                }
            }
        }

        /// <summary>
        /// Writes elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteValue(this.EmailAddress.Address, XmlElementNames.EmailAddress);
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
            jsonProperty.Add(XmlAttributeNames.Name, this.EmailAddress.Name);
            jsonProperty.Add(XmlAttributeNames.RoutingType, this.EmailAddress.RoutingType);

            if (this.EmailAddress.MailboxType.HasValue)
            {
                jsonProperty.Add(XmlAttributeNames.MailboxType, this.EmailAddress.MailboxType.Value);
            }

            jsonProperty.Add(XmlElementNames.EmailAddress, this.EmailAddress.Address);

            return jsonProperty;
        }

        /// <summary>
        /// Gets or sets the e-mail address of the entry.
        /// </summary>
        public EmailAddress EmailAddress
        {
            get
            {
                return this.emailAddress;
            }
            
            set
            {
                this.SetFieldValue<EmailAddress>(ref this.emailAddress, value);

                if (this.emailAddress != null)
                {
                    this.emailAddress.OnChange += this.EmailAddressChanged;
                }
            }
        }

        /// <summary>
        /// E-mail address was changed.
        /// </summary>
        /// <param name="complexProperty">Property that changed.</param>
        private void EmailAddressChanged(ComplexProperty complexProperty)
        {
            this.Changed();
        }
    }
}