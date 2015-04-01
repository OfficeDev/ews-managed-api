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
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents information about an attendee for which to request availability information.
    /// </summary>
    public sealed class AttendeeInfo : ISelfValidate
    {
        private string smtpAddress;
        private MeetingAttendeeType attendeeType = MeetingAttendeeType.Required;
        private bool excludeConflicts;

        /// <summary>
        /// Initializes a new instance of the <see cref="AttendeeInfo"/> class.
        /// </summary>
        public AttendeeInfo()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="AttendeeInfo"/> class.
        /// </summary>
        /// <param name="smtpAddress">The SMTP address of the attendee.</param>
        /// <param name="attendeeType">The yype of the attendee.</param>
        /// <param name="excludeConflicts">Indicates whether times when this attendee is not available should be returned.</param>
        public AttendeeInfo(
            string smtpAddress,
            MeetingAttendeeType attendeeType,
            bool excludeConflicts)
            : this()
        {
            this.smtpAddress = smtpAddress;
            this.attendeeType = attendeeType;
            this.excludeConflicts = excludeConflicts;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="AttendeeInfo"/> class.
        /// </summary>
        /// <param name="smtpAddress">The SMTP address of the attendee.</param>
        public AttendeeInfo(string smtpAddress)
            : this(smtpAddress, MeetingAttendeeType.Required, false)
        {
            this.smtpAddress = smtpAddress;
        }

        /// <summary>
        /// Defines an implicit conversion between a string representing an SMTP address and AttendeeInfo.
        /// </summary>
        /// <param name="smtpAddress">The SMTP address to convert to AttendeeInfo.</param>
        /// <returns>An AttendeeInfo initialized with the specified SMTP address.</returns>
        public static implicit operator AttendeeInfo(string smtpAddress)
        {
            return new AttendeeInfo(smtpAddress);
        }

        /// <summary>
        /// Writes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal void WriteToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.MailboxData);

            writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.Email);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Address, this.SmtpAddress);
            writer.WriteEndElement(); // Email

            writer.WriteElementValue(
                XmlNamespace.Types,
                XmlElementNames.AttendeeType,
                this.attendeeType);

            writer.WriteElementValue(
                XmlNamespace.Types,
                XmlElementNames.ExcludeConflicts,
                this.excludeConflicts);

            writer.WriteEndElement(); // MailboxData
        }

        /// <summary>
        /// Gets or sets the SMTP address of this attendee.
        /// </summary>
        public string SmtpAddress
        {
            get { return this.smtpAddress; }
            set { this.smtpAddress = value; }
        }

        /// <summary>
        /// Gets or sets the type of this attendee.
        /// </summary>
        public MeetingAttendeeType AttendeeType
        {
            get { return this.attendeeType; }
            set { this.attendeeType = value; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether times when this attendee is not available should be returned.
        /// </summary>
        public bool ExcludeConflicts
        {
            get { return this.excludeConflicts; }
            set { this.excludeConflicts = value; }
        }

        #region ISelfValidate Members

        /// <summary>
        /// Validates this instance.
        /// </summary>
        void ISelfValidate.Validate()
        {
            EwsUtilities.ValidateParam(this.smtpAddress, "SmtpAddress");
        }

        #endregion
    }
}