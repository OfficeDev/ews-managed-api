// ---------------------------------------------------------------------------
// <copyright file="GetRoomsRequest.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the GetPhoneCallRequest class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a GetRooms request.
    /// </summary>
    internal sealed class GetRoomsRequest : SimpleServiceRequestBase
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="GetRoomsRequest"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        internal GetRoomsRequest(ExchangeService service)
            : base(service)
        {
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.GetRoomsRequest;
        }

        /// <summary>
        /// Writes XML elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            this.RoomList.WriteToXml(writer, XmlNamespace.Messages, XmlElementNames.RoomList);
        }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.GetRoomsResponse;
        }

        /// <summary>
        /// Parses the response.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>Response object.</returns>
        internal override object ParseResponse(EwsServiceXmlReader reader)
        {
            GetRoomsResponse response = new GetRoomsResponse();
            response.LoadFromXml(reader, XmlElementNames.GetRoomsResponse);
            return response;
        }

        /// <summary>
        /// Gets the request version.
        /// </summary>
        /// <returns>Earliest Exchange version in which this request is supported.</returns>
        internal override ExchangeVersion GetMinimumRequiredServerVersion()
        {
            return ExchangeVersion.Exchange2010;
        }

        /// <summary>
        /// Executes this request.
        /// </summary>
        /// <returns>Service response.</returns>
        internal GetRoomsResponse Execute()
        {
            GetRoomsResponse serviceResponse = (GetRoomsResponse)this.InternalExecute();
            serviceResponse.ThrowIfNecessary();
            return serviceResponse;
        }

        /// <summary>
        /// Gets or sets the room list to retrieve rooms from.
        /// </summary>
        internal EmailAddress RoomList
        {
            get 
            {
                return this.roomList;
            }

            set 
            {
                this.roomList = value;
            }
        }

        private EmailAddress roomList;
    }
}