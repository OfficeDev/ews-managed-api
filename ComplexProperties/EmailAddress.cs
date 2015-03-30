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
    /// <summary>
    /// Represents an e-mail address.
    /// </summary>
    public class EmailAddress : ComplexProperty, ISearchStringProvider
    {
        /// <summary>
        /// SMTP routing type.
        /// </summary>
        internal const string SmtpRoutingType = "SMTP";

        /// <summary>
        /// Display name.
        /// </summary>
        private string name;

        /// <summary>
        /// Email address.
        /// </summary>
        private string address;

        /// <summary>
        /// Routing type.
        /// </summary>
        private string routingType;

        /// <summary>
        /// Mailbox type. 
        /// </summary>
        private MailboxType? mailboxType;

        /// <summary>
        /// ItemId - Contact or PDL.
        /// </summary>
        private ItemId id;

        /// <summary>
        /// Initializes a new instance of the <see cref="EmailAddress"/> class.
        /// </summary>
        public EmailAddress()
            : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="EmailAddress"/> class.
        /// </summary>
        /// <param name="smtpAddress">The SMTP address used to initialize the EmailAddress.</param>
        public EmailAddress(string smtpAddress)
            : this()
        {
            this.address = smtpAddress;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="EmailAddress"/> class.
        /// </summary>
        /// <param name="name">The name used to initialize the EmailAddress.</param>
        /// <param name="smtpAddress">The SMTP address used to initialize the EmailAddress.</param>
        public EmailAddress(string name, string smtpAddress)
            : this(smtpAddress)
        {
            this.name = name;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="EmailAddress"/> class.
        /// </summary>
        /// <param name="name">The name used to initialize the EmailAddress.</param>
        /// <param name="address">The address used to initialize the EmailAddress.</param>
        /// <param name="routingType">The routing type used to initialize the EmailAddress.</param>
        public EmailAddress(
            string name,
            string address,
            string routingType)
            : this(name, address)
        {
            this.routingType = routingType;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="EmailAddress"/> class.
        /// </summary>
        /// <param name="name">The name used to initialize the EmailAddress.</param>
        /// <param name="address">The address used to initialize the EmailAddress.</param>
        /// <param name="routingType">The routing type used to initialize the EmailAddress.</param>
        /// <param name="mailboxType">Mailbox type of the participant.</param>
        internal EmailAddress(
            string name,
            string address,
            string routingType,
            MailboxType mailboxType)
            : this(name, address, routingType)
        {
            this.mailboxType = mailboxType;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="EmailAddress"/> class.
        /// </summary>
        /// <param name="name">The name used to initialize the EmailAddress.</param>
        /// <param name="address">The address used to initialize the EmailAddress.</param>
        /// <param name="routingType">The routing type used to initialize the EmailAddress.</param>
        /// <param name="mailboxType">Mailbox type of the participant.</param>
        /// <param name="itemId">ItemId of a Contact or PDL.</param>
        internal EmailAddress(
            string name,
            string address,
            string routingType,
            MailboxType mailboxType,
            ItemId itemId)
            : this(name, address, routingType)
        {
            this.mailboxType = mailboxType;
            this.id = itemId;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="EmailAddress"/> class from another EmailAddress instance.
        /// </summary>
        /// <param name="mailbox">EMailAddress instance to copy.</param>
        internal EmailAddress(EmailAddress mailbox)
            : this()
        {
            EwsUtilities.ValidateParam(mailbox, "mailbox");

            this.Name = mailbox.Name;
            this.Address = mailbox.Address;
            this.RoutingType = mailbox.RoutingType;
            this.MailboxType = mailbox.MailboxType;
            this.Id = mailbox.Id;
        }

        /// <summary>
        /// Gets or sets the name associated with the e-mail address.
        /// </summary>
        public string Name
        {
            get
            {
                return this.name;
            }

            set
            {
                this.SetFieldValue<string>(ref this.name, value);
            }
        }

        /// <summary>
        /// Gets or sets the actual address associated with the e-mail address. The type of the Address property
        /// must match the specified routing type. If RoutingType is not set, Address is assumed to be an SMTP
        /// address.
        /// </summary>
        public string Address
        {
            get
            {
                return this.address;
            }

            set
            {
                this.SetFieldValue<string>(ref this.address, value);
            }
        }

        /// <summary>
        /// Gets or sets the routing type associated with the e-mail address. If RoutingType is not set,
        /// Address is assumed to be an SMTP address.
        /// </summary>
        public string RoutingType
        {
            get
            {
                return this.routingType;
            }

            set
            {
                this.SetFieldValue<string>(ref this.routingType, value);
            }
        }

        /// <summary>
        /// Gets or sets the type of the e-mail address.
        /// </summary>
        public MailboxType? MailboxType
        {
            get
            {
                return this.mailboxType;
            }

            set
            {
                this.SetFieldValue<MailboxType?>(ref this.mailboxType, value);
            }
        }

        /// <summary>
        /// Gets or sets the Id of the contact the e-mail address represents. When Id is specified, Address
        /// should be set to null.
        /// </summary>
        public ItemId Id
        {
            get
            {
                return this.id;
            }

            set
            {
                this.SetFieldValue<ItemId>(ref this.id, value);
            }
        }

        /// <summary>
        /// Defines an implicit conversion between a string representing an SMTP address and EmailAddress.
        /// </summary>
        /// <param name="smtpAddress">The SMTP address to convert to EmailAddress.</param>
        /// <returns>An EmailAddress initialized with the specified SMTP address.</returns>
        public static implicit operator EmailAddress(string smtpAddress)
        {
            return new EmailAddress(smtpAddress);
        }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            switch (reader.LocalName)
            {
                case XmlElementNames.Name:
                    this.name = reader.ReadElementValue();
                    return true;
                case XmlElementNames.EmailAddress:
                    this.address = reader.ReadElementValue();
                    return true;
                case XmlElementNames.RoutingType:
                    this.routingType = reader.ReadElementValue();
                    return true;
                case XmlElementNames.MailboxType:
                    this.mailboxType = reader.ReadElementValue<MailboxType>();
                    return true;
                case XmlElementNames.ItemId:
                    this.id = new ItemId();
                    this.id.LoadFromXml(reader, reader.LocalName);
                    return true;
                default:
                    return false;
            }
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
                    case XmlElementNames.Name:
                        this.name = jsonProperty.ReadAsString(key);
                        break;
                    case XmlElementNames.EmailAddress:
                        this.address = jsonProperty.ReadAsString(key);
                        break;
                    case XmlElementNames.RoutingType:
                        this.routingType = jsonProperty.ReadAsString(key);
                        break;
                    case XmlElementNames.MailboxType:
                        this.mailboxType = jsonProperty.ReadEnumValue<MailboxType>(key);
                        break;
                    case XmlElementNames.ItemId:
                        this.id = new ItemId();
                        this.id.LoadFromJson(jsonProperty.ReadAsJsonObject(key), service);
                        break;
                    default:
                        break;
                }
            }
        }

        /// <summary>
        /// Writes elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Name, this.Name);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.EmailAddress, this.Address);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.RoutingType, this.RoutingType);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.MailboxType, this.MailboxType);

            if (this.Id != null)
            {
                this.Id.WriteToXml(writer, XmlElementNames.ItemId);
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
            JsonObject jsonProperty = new JsonObject();

            jsonProperty.Add(XmlElementNames.Name, this.Name);
            jsonProperty.Add(XmlElementNames.EmailAddress, this.Address);
            jsonProperty.Add(XmlElementNames.RoutingType, this.RoutingType);
            jsonProperty.Add(XmlElementNames.MailboxType, this.MailboxType);

            if (this.Id != null)
            {
                jsonProperty.Add(XmlElementNames.ItemId, this.Id.InternalToJson(service));
            }

            return jsonProperty;
        }

        #region ISearchStringProvider methods
        /// <summary>
        /// Get a string representation for using this instance in a search filter.
        /// </summary>
        /// <returns>String representation of instance.</returns>
        string ISearchStringProvider.GetSearchString()
        {
            return this.Address;
        }
        #endregion

        #region Object method overrides
        /// <summary>
        /// Returns a <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </summary>
        /// <returns>
        /// A <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </returns>
        public override string ToString()
        {
            string addressPart;

            if (string.IsNullOrEmpty(this.Address))
            {
                return string.Empty;
            }

            if (!string.IsNullOrEmpty(this.RoutingType))
            {
                addressPart = this.RoutingType + ":" + this.Address;
            }
            else
            {
                addressPart = this.Address;
            }

            if (!string.IsNullOrEmpty(this.Name))
            {
                return this.Name + " <" + addressPart + ">";
            }
            else
            {
                return addressPart;
            }
        }
        #endregion
    }
}