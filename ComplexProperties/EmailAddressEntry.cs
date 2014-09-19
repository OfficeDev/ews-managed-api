// ---------------------------------------------------------------------------
// <copyright file="EmailAddressEntry.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

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