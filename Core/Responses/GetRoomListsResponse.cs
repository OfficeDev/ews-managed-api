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
// <summary>Defines the GetRoomListsResponse class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Represents the response to a GetRoomLists operation.
    /// </summary>
    internal sealed class GetRoomListsResponse : ServiceResponse
    {
        private EmailAddressCollection roomLists = new EmailAddressCollection();

        /// <summary>
        /// Initializes a new instance of the <see cref="GetRoomListsResponse"/> class.
        /// </summary>
        internal GetRoomListsResponse()
            : base()
        {
        }

        /// <summary>
        /// Gets all room list returned
        /// </summary>
        public EmailAddressCollection RoomLists
        {
            get { return this.roomLists; }
        }

        /// <summary>
        /// Reads response elements from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
        {
            this.RoomLists.Clear();
            base.ReadElementsFromXml(reader);

            reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.RoomLists);

            if (!reader.IsEmptyElement)
            {
                // Because we don't have an element for count of returned object,
                // we have to test the element to determine if it is return object or EndElement
                reader.Read();
                while (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.Address))
                {
                    EmailAddress emailAddress = new EmailAddress();
                    emailAddress.LoadFromXml(reader, XmlElementNames.Address);
                    this.RoomLists.Add(emailAddress);
                    reader.Read();
                }
                reader.EnsureCurrentNodeIsEndElement(XmlNamespace.Messages, XmlElementNames.RoomLists);
            }
            return;
        }
    }
}
