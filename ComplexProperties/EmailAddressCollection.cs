// ---------------------------------------------------------------------------
// <copyright file="EmailAddressCollection.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Implements an e-mail address collection.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// Represents a collection of e-mail addresses.
    /// </summary>
    public sealed class EmailAddressCollection : ComplexPropertyCollection<EmailAddress>
    {
        /// <summary>
        /// XML element name
        /// </summary>
        private string collectionItemXmlElementName;

        /// <summary>
        /// Initializes a new instance of the <see cref="EmailAddressCollection"/> class.
        /// </summary>
        /// <remarks>
        /// Note that XmlElementNames.Mailbox is the collection element name for ArrayOfRecipientsType, not ArrayOfEmailAddressesType.
        /// </remarks>
        internal EmailAddressCollection()
            : this(XmlElementNames.Mailbox)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="EmailAddressCollection"/> class.
        /// </summary>
        /// <param name="collectionItemXmlElementName">Name of the collection item XML element.</param>
        internal EmailAddressCollection(string collectionItemXmlElementName)
            : base()
        {
            this.collectionItemXmlElementName = collectionItemXmlElementName;
        }

        /// <summary>
        /// Adds an e-mail address to the collection.
        /// </summary>
        /// <param name="emailAddress">The e-mail address to add.</param>
        public void Add(EmailAddress emailAddress)
        {
            this.InternalAdd(emailAddress);
        }

        /// <summary>
        /// Adds multiple e-mail addresses to the collection.
        /// </summary>
        /// <param name="emailAddresses">The e-mail addresses to add.</param>
        public void AddRange(IEnumerable<EmailAddress> emailAddresses)
        {
            foreach (EmailAddress emailAddress in emailAddresses)
            {
                this.Add(emailAddress);
            }
        }

        /// <summary>
        /// Adds an e-mail address to the collection.
        /// </summary>
        /// <param name="smtpAddress">The SMTP address used to initialize the e-mail address.</param>
        /// <returns>An EmailAddress object initialized with the provided SMTP address.</returns>
        public EmailAddress Add(string smtpAddress)
        {
            EmailAddress emailAddress = new EmailAddress(smtpAddress);

            this.Add(emailAddress);

            return emailAddress;
        }

        /// <summary>
        /// Adds multiple e-mail addresses to the collection.
        /// </summary>
        /// <param name="smtpAddresses">The SMTP addresses used to initialize the e-mail addresses.</param>
        public void AddRange(IEnumerable<string> smtpAddresses)
        {
            foreach (string smtpAddress in smtpAddresses)
            {
                this.Add(smtpAddress);
            }
        }

        /// <summary>
        /// Adds an e-mail address to the collection.
        /// </summary>
        /// <param name="name">The name used to initialize the e-mail address.</param>
        /// <param name="smtpAddress">The SMTP address used to initialize the e-mail address.</param>
        /// <returns>An EmailAddress object initialized with the provided SMTP address.</returns>
        public EmailAddress Add(string name, string smtpAddress)
        {
            EmailAddress emailAddress = new EmailAddress(name, smtpAddress);

            this.Add(emailAddress);

            return emailAddress;
        }

        /// <summary>
        /// Clears the collection.
        /// </summary>
        public void Clear()
        {
            this.InternalClear();
        }

        /// <summary>
        /// Removes an e-mail address from the collection.
        /// </summary>
        /// <param name="index">The index of the e-mail address to remove.</param>
        public void RemoveAt(int index)
        {
            if (index < 0 || index >= this.Count)
            {
                throw new ArgumentOutOfRangeException("index", Strings.IndexIsOutOfRange);
            }

            this.InternalRemoveAt(index);
        }

        /// <summary>
        /// Removes an e-mail address from the collection.
        /// </summary>
        /// <param name="emailAddress">The e-mail address to remove.</param>
        /// <returns>True if the email address was successfully removed from the collection, false otherwise.</returns>
        public bool Remove(EmailAddress emailAddress)
        {
            EwsUtilities.ValidateParam(emailAddress, "emailAddress");

            return this.InternalRemove(emailAddress);
        }

        /// <summary>
        /// Creates an EmailAddress object from an XML element name.
        /// </summary>
        /// <param name="xmlElementName">The XML element name from which to create the e-mail address.</param>
        /// <returns>An EmailAddress object.</returns>
        internal override EmailAddress CreateComplexProperty(string xmlElementName)
        {
            if (xmlElementName == this.collectionItemXmlElementName)
            {
                return new EmailAddress();
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Creates the default complex property.
        /// </summary>
        /// <returns></returns>
        internal override EmailAddress CreateDefaultComplexProperty()
        {
            return new EmailAddress();
        }

        /// <summary>
        /// Retrieves the XML element name corresponding to the provided EmailAddress object.
        /// </summary>
        /// <param name="emailAddress">The EmailAddress object from which to determine the XML element name.</param>
        /// <returns>The XML element name corresponding to the provided EmailAddress object.</returns>
        internal override string GetCollectionItemXmlElementName(EmailAddress emailAddress)
        {
            return this.collectionItemXmlElementName;
        }

        /// <summary>
        /// Determine whether we should write collection to XML or not.
        /// </summary>
        /// <returns>Always true, even if the collection is empty.</returns>
        internal override bool ShouldWriteToRequest()
        {
            return true;
        }
    }
}
