// ---------------------------------------------------------------------------
// <copyright file="MeetingSuggestion.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the MeetingSuggestion class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.IO;

    /// <summary>
    /// Represents an MeetingSuggestion object.
    /// </summary>
    public sealed class MeetingSuggestion : ExtractedEntity
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="MeetingSuggestion"/> class.
        /// </summary>
        internal MeetingSuggestion()
            : base()
        {
        }

        /// <summary>
        /// Gets the meeting suggestion Attendees.
        /// </summary>
        public EmailUserEntityCollection Attendees { get; internal set; }

        /// <summary>
        /// Gets the meeting suggestion Location.
        /// </summary>
        public string Location { get; internal set; }

        /// <summary>
        /// Gets the meeting suggestion Subject.
        /// </summary>
        public string Subject { get; internal set; }

        /// <summary>
        /// Gets the meeting suggestion MeetingString.
        /// </summary>
        public string MeetingString { get; internal set; }

        /// <summary>
        /// Gets the meeting suggestion StartTime.
        /// </summary>
        public DateTime? StartTime { get; internal set; }

        /// <summary>
        /// Gets the meeting suggestion EndTime.
        /// </summary>
        public DateTime? EndTime { get; internal set; }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            switch (reader.LocalName)
            {
                case XmlElementNames.NlgAttendees:
                    this.Attendees = new EmailUserEntityCollection();
                    this.Attendees.LoadFromXml(reader, XmlNamespace.Types, XmlElementNames.NlgAttendees);
                    return true;

                case XmlElementNames.NlgLocation:
                    this.Location = reader.ReadElementValue();
                    return true;

                case XmlElementNames.NlgSubject:
                    this.Subject = reader.ReadElementValue();
                    return true;

                case XmlElementNames.NlgMeetingString:
                    this.MeetingString = reader.ReadElementValue();
                    return true;

                case XmlElementNames.NlgStartTime:
                    this.StartTime = reader.ReadElementValueAsDateTime();
                    return true;

                case XmlElementNames.NlgEndTime:
                    this.EndTime = reader.ReadElementValueAsDateTime();
                    return true;
                
                default:
                    return base.TryReadElementFromXml(reader);
            }
        }
    }
}
