// ---------------------------------------------------------------------------
// <copyright file="UpdateDelegateRequest.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the UpdateDelegateRequest class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System.Collections.Generic;

    /// <summary>
    /// Represents an UpdateDelegate request.
    /// </summary>
    internal class UpdateDelegateRequest : DelegateManagementRequestBase<DelegateManagementResponse>
    {
        private List<DelegateUser> delegateUsers = new List<DelegateUser>();
        private MeetingRequestsDeliveryScope? meetingRequestsDeliveryScope;

        /// <summary>
        /// Initializes a new instance of the <see cref="UpdateDelegateRequest"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        internal UpdateDelegateRequest(ExchangeService service)
            : base(service)
        {
        }

        /// <summary>
        /// Validate request..
        /// </summary>
        internal override void Validate()
        {
            base.Validate();
            EwsUtilities.ValidateParamCollection(this.DelegateUsers, "DelegateUsers");
            
            foreach (DelegateUser delegateUser in this.DelegateUsers)
            {
                delegateUser.ValidateUpdateDelegate();
            }
        }

        /// <summary>
        /// Writes XML elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            base.WriteElementsToXml(writer);

            writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.DelegateUsers);

            foreach (DelegateUser delegateUser in this.DelegateUsers)
            {
                delegateUser.WriteToXml(writer, XmlElementNames.DelegateUser);
            }

            writer.WriteEndElement(); // DelegateUsers

            if (this.MeetingRequestsDeliveryScope.HasValue)
            {
                writer.WriteElementValue(
                XmlNamespace.Messages,
                XmlElementNames.DeliverMeetingRequests,
                this.MeetingRequestsDeliveryScope.Value);
            }
        }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.UpdateDelegateResponse;
        }

        /// <summary>
        /// Creates the response.
        /// </summary>
        /// <returns>Response object.</returns>
        internal override DelegateManagementResponse CreateResponse()
        {
            return new DelegateManagementResponse(true, this.delegateUsers);
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.UpdateDelegate;
        }

        /// <summary>
        /// Gets the request version.
        /// </summary>
        /// <returns>Earliest Exchange version in which this request is supported.</returns>
        internal override ExchangeVersion GetMinimumRequiredServerVersion()
        {
            return ExchangeVersion.Exchange2007_SP1;
        }

        /// <summary>
        /// Gets or sets the meeting requests delivery scope.
        /// </summary>
        /// <value>The meeting requests delivery scope.</value>
        public MeetingRequestsDeliveryScope? MeetingRequestsDeliveryScope
        {
            get { return this.meetingRequestsDeliveryScope; }
            set { this.meetingRequestsDeliveryScope = value; }
        }

        /// <summary>
        /// Gets the delegate users.
        /// </summary>
        /// <value>The delegate users.</value>
        public List<DelegateUser> DelegateUsers
        {
            get { return this.delegateUsers; }
        }
    }
}
