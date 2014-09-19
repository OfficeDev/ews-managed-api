// ---------------------------------------------------------------------------
// <copyright file="GetRoomListsResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

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
