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
// <summary>Defines the EntityExtractionResult class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.IO;

    /// <summary>
    /// Represents an EntityExtractionResult object.
    /// </summary>
    public sealed class EntityExtractionResult : ComplexProperty
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="EntityExtractionResult"/> class.
        /// </summary>
        internal EntityExtractionResult()
            : base()
        {
            this.Namespace = XmlNamespace.Types;
        }

        /// <summary>
        /// Gets the extracted Addresses.
        /// </summary>
        public AddressEntityCollection Addresses { get; internal set; }

        /// <summary>
        /// Gets the extracted MeetingSuggestions.
        /// </summary>
        public MeetingSuggestionCollection MeetingSuggestions { get; internal set; }

        /// <summary>
        /// Gets the extracted TaskSuggestions.
        /// </summary>
        public TaskSuggestionCollection TaskSuggestions { get; internal set; }

        /// <summary>
        /// Gets the extracted EmailAddresses.
        /// </summary>
        public EmailAddressEntityCollection EmailAddresses { get; internal set; }

        /// <summary>
        /// Gets the extracted Contacts.
        /// </summary>
        public ContactEntityCollection Contacts { get; internal set; }

        /// <summary>
        /// Gets the extracted Urls.
        /// </summary>
        public UrlEntityCollection Urls { get; internal set; }

        /// <summary>
        /// Gets the extracted PhoneNumbers.
        /// </summary>
        public PhoneEntityCollection PhoneNumbers { get; internal set; }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            switch (reader.LocalName)
            {
                case XmlElementNames.NlgAddresses:
                    this.Addresses = new AddressEntityCollection();
                    this.Addresses.LoadFromXml(reader, XmlNamespace.Types, XmlElementNames.NlgAddresses);
                    return true;

                case XmlElementNames.NlgMeetingSuggestions:
                    this.MeetingSuggestions = new MeetingSuggestionCollection();
                    this.MeetingSuggestions.LoadFromXml(reader, XmlNamespace.Types, XmlElementNames.NlgMeetingSuggestions);
                    return true;

                case XmlElementNames.NlgTaskSuggestions:
                    this.TaskSuggestions = new TaskSuggestionCollection();
                    this.TaskSuggestions.LoadFromXml(reader, XmlNamespace.Types, XmlElementNames.NlgTaskSuggestions);
                    return true;

                case XmlElementNames.NlgEmailAddresses:
                    this.EmailAddresses = new EmailAddressEntityCollection();
                    this.EmailAddresses.LoadFromXml(reader, XmlNamespace.Types, XmlElementNames.NlgEmailAddresses);
                    return true;

                case XmlElementNames.NlgContacts:
                    this.Contacts = new ContactEntityCollection();
                    this.Contacts.LoadFromXml(reader, XmlNamespace.Types, XmlElementNames.NlgContacts);
                    return true;

                case XmlElementNames.NlgUrls:
                    this.Urls = new UrlEntityCollection();
                    this.Urls.LoadFromXml(reader, XmlNamespace.Types, XmlElementNames.NlgUrls);
                    return true;

                case XmlElementNames.NlgPhoneNumbers:
                    this.PhoneNumbers = new PhoneEntityCollection();
                    this.PhoneNumbers.LoadFromXml(reader, XmlNamespace.Types, XmlElementNames.NlgPhoneNumbers);
                    return true;
                
                default:
                    return base.TryReadElementFromXml(reader);
            }
        }
    }
}
