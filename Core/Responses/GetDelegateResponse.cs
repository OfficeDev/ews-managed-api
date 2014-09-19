// ---------------------------------------------------------------------------
// <copyright file="GetDelegateResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the GetDelegateResponse class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Represents the response to a delegate user retrieval operation.
    /// </summary>
    internal sealed class GetDelegateResponse : DelegateManagementResponse
    {
        private MeetingRequestsDeliveryScope meetingRequestsDeliveryScope = MeetingRequestsDeliveryScope.NoForward;

        /// <summary>
        /// Initializes a new instance of the <see cref="GetDelegateResponse"/> class.
        /// </summary>
        /// <param name="readDelegateUsers">if set to <c>true</c> [read delegate users].</param>
        internal GetDelegateResponse(bool readDelegateUsers)
            : base(readDelegateUsers, null)
        {
        }

        /// <summary>
        /// Reads response elements from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
        {
            base.ReadElementsFromXml(reader);

            if (this.ErrorCode == ServiceError.NoError)
            {
                // If there were no response messages, the reader will already be on the
                // DeliverMeetingRequests start element, so we don't need to read it.
                if (this.DelegateUserResponses.Count > 0)
                {
                    reader.Read();
                }

                // Make sure that we're at the DeliverMeetingRequests element before trying to read the value.
                // In error cases, the element may not have been returned.
                if (reader.IsStartElement(XmlNamespace.Messages, XmlElementNames.DeliverMeetingRequests))
                {
                    this.meetingRequestsDeliveryScope = reader.ReadElementValue<MeetingRequestsDeliveryScope>();
                }
            }
        }

        /// <summary>
        /// Gets a value indicating if and how meeting requests are delivered to delegates.
        /// </summary>
        internal MeetingRequestsDeliveryScope MeetingRequestsDeliveryScope
        {
            get { return this.meetingRequestsDeliveryScope; }
        }
    }
}
