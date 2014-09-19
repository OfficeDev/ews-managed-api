// ---------------------------------------------------------------------------
// <copyright file="GetRoomsResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the GetPhoneCallResponse class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Text;

    /// <summary>
    /// Represents the response to a GetRooms operation.
    /// </summary>
    internal sealed class GetRoomsResponse : ServiceResponse
    {
        private Collection<EmailAddress> rooms = new Collection<EmailAddress>();

        /// <summary>
        /// Initializes a new instance of the <see cref="GetRoomsResponse"/> class.
        /// </summary>
        internal GetRoomsResponse()
            : base()
        {
        }

        /// <summary>
        /// Gets collection for all rooms returned
        /// </summary>
        public Collection<EmailAddress> Rooms
        {
            get { return this.rooms; }
        }

        /// <summary>
        /// Reads response elements from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
        {
            this.Rooms.Clear();
            base.ReadElementsFromXml(reader);

            reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.Rooms);

            if (!reader.IsEmptyElement)
            {
                // Because we don't have an element for count of returned object,
                // we have to test the element to determine if it is StartElement of return object or EndElement
                reader.Read();
                while (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.Room))
                {
                    reader.Read(); // skip the start <Room>

                    EmailAddress emailAddress = new EmailAddress();
                    emailAddress.LoadFromXml(reader, XmlElementNames.RoomId);
                    this.Rooms.Add(emailAddress);

                    reader.ReadEndElement(XmlNamespace.Types, XmlElementNames.Room);
                    reader.Read();
                }

                reader.EnsureCurrentNodeIsEndElement(XmlNamespace.Messages, XmlElementNames.Rooms);
            }
        }
    }
}