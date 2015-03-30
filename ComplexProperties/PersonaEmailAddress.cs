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
    public sealed class PersonaEmailAddress : ComplexProperty, ISearchStringProvider
    {
        /// <summary>
        /// Creates a new instance of the <see cref="PersonaEmailAddress"/> class.
        /// </summary>
        public PersonaEmailAddress()
            : base()
        {
            _emailAddress = new EmailAddress();
        }

        /// <summary>
        /// Creates a new instance of the <see cref="EmailAddress"/> class.
        /// </summary>
        /// <param name="smtpAddress">The SMTP address used to initialize the PersonaEmailAddress.</param>
        public PersonaEmailAddress(string smtpAddress)
            : this()
        {
            EwsUtilities.ValidateParam(smtpAddress, "smtpAddress");
            this.Address = smtpAddress;
        }

        /// <summary>
        /// Creates a new instance of the <see cref="PersonaEmailAddress"/> class.
        /// </summary>
        /// <param name="name">The name used to initialize the PersonaEmailAddress.</param>
        /// <param name="smtpAddress">The SMTP address used to initialize the PersonaEmailAddress.</param>
        public PersonaEmailAddress(string name, string smtpAddress)
            : this(smtpAddress)
        {
            EwsUtilities.ValidateParam(name, "name");
            this.Name = name;
        }

        /// <summary>
        /// Name accessors
        /// </summary>
        public string Name
        {
            get
            {
                return _emailAddress.Name;
            }

            set
            {
                _emailAddress.Name = value;
            }
        }

        /// <summary>
        /// Email address accessors. The type of the Address property must match the specified routing type.
        /// If RoutingType is not set, Address is assumed to be an SMTP address.
        /// </summary>
        public string Address
        {
            get
            {
                return _emailAddress.Address;
            }

            set
            {
                _emailAddress.Address = value;
            }
        }

        /// <summary>
        /// Routing type accessors. If RoutingType is not set, Address is assumed to be an SMTP address.
        /// </summary>
        public string RoutingType
        {
            get
            {
                return _emailAddress.RoutingType;
            }

            set
            {
                _emailAddress.RoutingType = value;
            }
        }

        /// <summary>
        /// Mailbox type accessors
        /// </summary>
        public MailboxType? MailboxType
        {
            get
            {
                return _emailAddress.MailboxType;
            }

            set
            {
                _emailAddress.MailboxType = value;
            }
        }

        /// <summary>
        /// PersonaEmailAddress Id accessors
        /// </summary>
        public ItemId Id
        {
            get
            {
                return _emailAddress.Id;
            }

            set
            {
                _emailAddress.Id = value;
            }
        }

        /// <summary>
        /// Original display name accessors
        /// </summary>
        public string OriginalDisplayName { get; set; }

        /// <summary>
        /// Email address details
        /// </summary>
        private EmailAddress _emailAddress;

        /// <summary>
        /// Defines an implicit conversion from a string representing an SMTP address to PeronaEmailAddress.
        /// </summary>
        /// <param name="smtpAddress">The SMTP address to convert to EmailAddress.</param>
        /// <returns>An EmailAddress initialized with the specified SMTP address.</returns>
        public static implicit operator PersonaEmailAddress(string smtpAddress)
        {
            return new PersonaEmailAddress(smtpAddress);
        }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">XML reader</param>
        /// <returns>Whether the element was read</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            while (true)
            {
                switch (reader.LocalName)
                {
                    case XmlElementNames.Name:
                        this.Name = reader.ReadElementValue();
                        break;
                    case XmlElementNames.EmailAddress:
                        this.Address = reader.ReadElementValue();

                        // Process the next node before returning. Otherwise, the current </EmailAddress> node
                        // makes ComplexProperty.InternalLoadFromXml think that this ends the outer <EmailAddress>
                        // node, causing the remaining children of the outer EmailAddress node to be skipped.
                        reader.Read();
                        if (reader.NodeType == System.Xml.XmlNodeType.Element)
                        {
                            continue;
                        }
                        break;
                    case XmlElementNames.RoutingType:
                        this.RoutingType = reader.ReadElementValue();
                        break;
                    case XmlElementNames.MailboxType:
                        this.MailboxType = reader.ReadElementValue<MailboxType>();
                        break;
                    case XmlElementNames.ItemId:
                        this.Id = new ItemId();
                        this.Id.LoadFromXml(reader, reader.LocalName);
                        break;
                    case XmlElementNames.OriginalDisplayName:
                        this.OriginalDisplayName = reader.ReadElementValue();
                        break;
                    default:
                        return false;
                }

                return true;
            }
        }

        /// <summary>
        /// Writes elements to XML.
        /// </summary>
        /// <param name="writer">XML writer</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Name, this.Name);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.EmailAddress, this.Address);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.RoutingType, this.RoutingType);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.MailboxType, this.MailboxType);

            if (!string.IsNullOrEmpty(this.OriginalDisplayName))
            {
                writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.OriginalDisplayName, this.OriginalDisplayName);
            }

            if (this.Id != null)
            {
                this.Id.WriteToXml(writer, XmlElementNames.ItemId);
            }
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