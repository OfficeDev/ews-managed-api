// ---------------------------------------------------------------------------
// <copyright file="PersonaEmailAddressCollection.cs" company="Microsoft">
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
    /// Represents a collection of persona e-mail addresses.
    /// </summary>
    public sealed class PersonaEmailAddressCollection : ComplexPropertyCollection<PersonaEmailAddress>
    {
        /// <summary>
        /// XML element name
        /// </summary>
        private readonly string collectionItemXmlElementName;

        /// <summary>
        /// Creates a new instance of the <see cref="PersonaEmailAddressCollection"/> class.
        /// </summary>
        /// <remarks>
        /// MSDN example incorrectly shows root element as EmailAddress. In fact, it is Address.
        /// </remarks>
        internal PersonaEmailAddressCollection()
            : this(XmlElementNames.Address)
        {
        }

        /// <summary>
        /// Creates a new instance of the <see cref="PersonaEmailAddressCollection"/> class.
        /// </summary>
        /// <param name="collectionItemXmlElementName">Name of the collection item XML element.</param>
        internal PersonaEmailAddressCollection(string collectionItemXmlElementName)
            : base()
        {
            this.collectionItemXmlElementName = collectionItemXmlElementName;
        }

        /// <summary>
        /// Adds a persona e-mail address to the collection.
        /// </summary>
        /// <param name="emailAddress">The persona e-mail address to add.</param>
        public void Add(PersonaEmailAddress emailAddress)
        {
            this.InternalAdd(emailAddress);
        }

        /// <summary>
        /// Adds multiple persona e-mail addresses to the collection.
        /// </summary>
        /// <param name="emailAddresses">The collection of persona e-mail addresses to add.</param>
        public void AddRange(IEnumerable<PersonaEmailAddress> emailAddresses)
        {
            if (emailAddresses != null)
            {
                foreach (PersonaEmailAddress emailAddress in emailAddresses)
                {
                    this.Add(emailAddress);
                }
            }
        }

        /// <summary>
        /// Adds a persona e-mail address to the collection.
        /// </summary>
        /// <param name="smtpAddress">The SMTP address used to initialize the persona e-mail address.</param>
        /// <returns>An PersonaEmailAddress object initialized with the provided SMTP address.</returns>
        public PersonaEmailAddress Add(string smtpAddress)
        {
            PersonaEmailAddress emailAddress = new PersonaEmailAddress(smtpAddress);

            this.Add(emailAddress);

            return emailAddress;
        }

        /// <summary>
        /// Adds multiple e-mail addresses to the collection.
        /// </summary>
        /// <param name="smtpAddresses">The SMTP addresses to be added as persona email addresses</param>
        public void AddRange(IEnumerable<string> smtpAddresses)
        {
            if (smtpAddresses != null)
            {
                foreach (string smtpAddress in smtpAddresses)
                {
                    this.Add(smtpAddress);
                }
            }
        }

        /// <summary>
        /// Adds an e-mail address to the collection.
        /// </summary>
        /// <param name="name">The name used to initialize the persona e-mail address.</param>
        /// <param name="smtpAddress">The SMTP address used to initialize the persona e-mail address.</param>
        /// <returns>An PersonaEmailAddress object initialized with the provided SMTP address.</returns>
        public PersonaEmailAddress Add(string name, string smtpAddress)
        {
            PersonaEmailAddress emailAddress = new PersonaEmailAddress(name, smtpAddress);

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
        /// Removes a persona e-mail address from the collection.
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
        /// Removes a persona e-mail address from the collection.
        /// </summary>
        /// <param name="personaEmailAddress">The e-mail address to remove.</param>
        /// <returns>Whether removed from the collection</returns>
        public bool Remove(PersonaEmailAddress personaEmailAddress)
        {
            EwsUtilities.ValidateParam(personaEmailAddress, "personaEmailAddress");

            return this.InternalRemove(personaEmailAddress);
        }

        /// <summary>
        /// Creates a PersonaEmailAddress object from an XML element name.
        /// </summary>
        /// <param name="xmlElementName">The XML element name from which to create the persona e-mail address.</param>
        /// <returns>A PersonaEmailAddress object.</returns>
        internal override PersonaEmailAddress CreateComplexProperty(string xmlElementName)
        {
            if (xmlElementName == this.collectionItemXmlElementName)
            {
                return new PersonaEmailAddress();
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
        internal override PersonaEmailAddress CreateDefaultComplexProperty()
        {
            return new PersonaEmailAddress();
        }

        /// <summary>
        /// Retrieves the XML element name corresponding to the provided PersonaEmailAddress object.
        /// </summary>
        /// <param name="personaEmailAddress">The PersonaEmailAddress object from which to determine the XML element name.</param>
        /// <returns>The XML element name corresponding to the provided PersonaEmailAddress object.</returns>
        internal override string GetCollectionItemXmlElementName(PersonaEmailAddress personaEmailAddress)
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
