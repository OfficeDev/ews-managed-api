// ---------------------------------------------------------------------------
// <copyright file="Attendee.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the Attendee class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;
    using System.Xml;

    /// <summary>
    /// Represents an attendee to a meeting.
    /// </summary>
    public sealed class Attendee : EmailAddress
    {
        private MeetingResponseType? responseType;
        private DateTime? lastResponseTime;

        /// <summary>
        /// Initializes a new instance of the <see cref="Attendee"/> class.
        /// </summary>
        public Attendee()
            : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Attendee"/> class.
        /// </summary>
        /// <param name="smtpAddress">The SMTP address used to initialize the Attendee.</param>
        public Attendee(string smtpAddress)
            : base(smtpAddress)
        {
            EwsUtilities.ValidateParam(smtpAddress, "smtpAddress");
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Attendee"/> class.
        /// </summary>
        /// <param name="name">The name used to initialize the Attendee.</param>
        /// <param name="smtpAddress">The SMTP address used to initialize the Attendee.</param>
        public Attendee(string name, string smtpAddress)
            : base(name, smtpAddress)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Attendee"/> class.
        /// </summary>
        /// <param name="name">The name used to initialize the Attendee.</param>
        /// <param name="smtpAddress">The SMTP address used to initialize the Attendee.</param>
        /// <param name="routingType">The routing type used to initialize the Attendee.</param>
        public Attendee(
            string name,
            string smtpAddress,
            string routingType)
            : base(name, smtpAddress, routingType)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Attendee"/> class from an EmailAddress.
        /// </summary>
        /// <param name="mailbox">The mailbox used to initialize the Attendee.</param>
        public Attendee(EmailAddress mailbox)
            : base(mailbox)
        {
        }

        /// <summary>
        /// Gets the type of response the attendee gave to the meeting invitation it received.
        /// </summary>
        public MeetingResponseType? ResponseType
        {
            get { return this.responseType; }
        }

        /// <summary>
        /// Gets the date and time when the attendee last responded to a meeting invitation or update.
        /// </summary>
        public DateTime? LastResponseTime
        {
            get { return this.lastResponseTime; }
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
                case XmlElementNames.Mailbox:
                    this.LoadFromXml(reader, reader.LocalName);
                    return true;
                case XmlElementNames.ResponseType:
                    this.responseType = reader.ReadElementValue<MeetingResponseType>();
                    return true;
                case XmlElementNames.LastResponseTime:
                    this.lastResponseTime = reader.ReadElementValueAsDateTime();
                    return true;
                default:
                    return base.TryReadElementFromXml(reader);
            }
        }
    
        /// <summary>
        /// Writes the elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteStartElement(this.Namespace, XmlElementNames.Mailbox);
            base.WriteElementsToXml(writer);
            writer.WriteEndElement();
        }
    }
}
