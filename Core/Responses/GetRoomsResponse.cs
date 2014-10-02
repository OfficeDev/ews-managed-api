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