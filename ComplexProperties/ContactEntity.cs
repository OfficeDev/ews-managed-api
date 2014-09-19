// ---------------------------------------------------------------------------
// <copyright file="ContactEntity.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ContactEntity class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.IO;

    /// <summary>
    /// Represents an ContactEntity object.
    /// </summary>
    public sealed class ContactEntity : ExtractedEntity
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ContactEntity"/> class.
        /// </summary>
        internal ContactEntity()
            : base()
        {
        }

        /// <summary>
        /// Gets the contact entity PersonName.
        /// </summary>
        public string PersonName { get; internal set; }

        /// <summary>
        /// Gets the contact entity BusinessName.
        /// </summary>
        public string BusinessName { get; internal set; }

        /// <summary>
        /// Gets the contact entity PhoneNumbers.
        /// </summary>
        public ContactPhoneEntityCollection PhoneNumbers { get; internal set; }

        /// <summary>
        /// Gets the contact entity Urls.
        /// </summary>
        public StringList Urls { get; internal set; }

        /// <summary>
        /// Gets the contact entity EmailAddresses.
        /// </summary>
        public StringList EmailAddresses { get; internal set; }

        /// <summary>
        /// Gets the contact entity Addresses.
        /// </summary>
        public StringList Addresses { get; internal set; }

        /// <summary>
        /// Gets the contact entity ContactString.
        /// </summary>
        public string ContactString { get; internal set; }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            switch (reader.LocalName)
            {
                case XmlElementNames.NlgPersonName:
                    this.PersonName = reader.ReadElementValue();
                    return true;

                case XmlElementNames.NlgBusinessName:
                    this.BusinessName = reader.ReadElementValue();
                    return true;

                case XmlElementNames.NlgPhoneNumbers:
                    this.PhoneNumbers = new ContactPhoneEntityCollection();
                    this.PhoneNumbers.LoadFromXml(reader, XmlNamespace.Types, XmlElementNames.NlgPhoneNumbers);
                    return true;

                case XmlElementNames.NlgUrls:
                    this.Urls = new StringList(XmlElementNames.NlgUrl);
                    this.Urls.LoadFromXml(reader, XmlNamespace.Types, XmlElementNames.NlgUrls);
                    return true;

                case XmlElementNames.NlgEmailAddresses:
                    this.EmailAddresses = new StringList(XmlElementNames.NlgEmailAddress);
                    this.EmailAddresses.LoadFromXml(reader, XmlNamespace.Types, XmlElementNames.NlgEmailAddresses);
                    return true;

                case XmlElementNames.NlgAddresses:
                    this.Addresses = new StringList(XmlElementNames.NlgAddress);
                    this.Addresses.LoadFromXml(reader, XmlNamespace.Types, XmlElementNames.NlgAddresses);
                    return true;

                case XmlElementNames.NlgContactString:
                    this.ContactString = reader.ReadElementValue();
                    return true;

                default:
                    return base.TryReadElementFromXml(reader);
            }
        }
    }
}
