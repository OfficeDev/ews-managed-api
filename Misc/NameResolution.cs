// ---------------------------------------------------------------------------
// <copyright file="NameResolution.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the NameResolution class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a suggested name resolution.
    /// </summary>
    public sealed class NameResolution
    {
        private NameResolutionCollection owner;
        private EmailAddress mailbox = new EmailAddress();
        private Contact contact;

        /// <summary>
        /// Initializes a new instance of the <see cref="NameResolution"/> class.
        /// </summary>
        /// <param name="owner">The owner.</param>
        internal NameResolution(NameResolutionCollection owner)
        {
            EwsUtilities.Assert(
                owner != null,
                "NameResolution.ctor",
                "owner is null.");

            this.owner = owner;
        }

        /// <summary>
        /// Loads from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal void LoadFromXml(EwsServiceXmlReader reader)
        {
            reader.ReadStartElement(XmlNamespace.Types, XmlElementNames.Resolution);

            reader.ReadStartElement(XmlNamespace.Types, XmlElementNames.Mailbox);
            this.mailbox.LoadFromXml(reader, XmlElementNames.Mailbox);

            reader.Read();
            if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.Contact))
            {
                this.contact = new Contact(this.owner.Session);

                // Contacts returned by ResolveNames should behave like Contact.Load with FirstClassPropertySet specified.
                this.contact.LoadFromXml(
                                    reader,
                                    true,                               /* clearPropertyBag */
                                    PropertySet.FirstClassProperties,
                                    false);                             /* summaryPropertiesOnly */

                reader.ReadEndElement(XmlNamespace.Types, XmlElementNames.Resolution);
            }
            else
            {
                reader.EnsureCurrentNodeIsEndElement(XmlNamespace.Types, XmlElementNames.Resolution);
            }
        }

        /// <summary>
        /// Loads from json.
        /// </summary>
        /// <param name="jsonProperty">The json property.</param>
        /// <param name="service">The service.</param>
        internal void LoadFromJson(JsonObject jsonProperty, ExchangeService service)
        {
            foreach (string key in jsonProperty.Keys)
            {
                switch (key)
                {
                    case XmlElementNames.Mailbox:
                        this.mailbox.LoadFromJson(jsonProperty.ReadAsJsonObject(key), service);
                        break;
                    case XmlElementNames.Contact:
                        this.contact = new Contact(this.owner.Session);
                        this.contact.LoadFromJson(jsonProperty.ReadAsJsonObject(key), service, true, PropertySet.FirstClassProperties, false);
                        break;
                    default:
                        break;
                }
            }
        }

        /// <summary>
        /// Gets the mailbox of the suggested resolved name.
        /// </summary>
        public EmailAddress Mailbox
        {
            get { return this.mailbox; }
        }

        /// <summary>
        /// Gets the contact information of the suggested resolved name. This property is only available when
        /// ResolveName is called with returnContactDetails = true.
        /// </summary>
        public Contact Contact
        {
            get { return this.contact; }
        }
    }
}
