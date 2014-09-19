// ---------------------------------------------------------------------------
// <copyright file="AddDelegateRequest.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the AddDelegateRequest class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System.Collections.Generic;

    /// <summary>
    /// Represents an AddDelegate request.
    /// </summary>
    internal class AddDelegateRequest : DelegateManagementRequestBase<DelegateManagementResponse>
    {
        private List<DelegateUser> delegateUsers = new List<DelegateUser>();
        private MeetingRequestsDeliveryScope? meetingRequestsDeliveryScope;

        /// <summary>
        /// Initializes a new instance of the <see cref="AddDelegateRequest"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        internal AddDelegateRequest(ExchangeService service)
            : base(service)
        {
        }

        /// <summary>
        /// Validate request.
        /// </summary>
        internal override void Validate()
        {
            base.Validate();
            EwsUtilities.ValidateParamCollection(this.DelegateUsers, "DelegateUsers");

            foreach (DelegateUser delegateUser in this.DelegateUsers)
            {
                delegateUser.ValidateUpdateDelegate();
            }

            if (this.MeetingRequestsDeliveryScope.HasValue)
            {
                EwsUtilities.ValidateEnumVersionValue(this.MeetingRequestsDeliveryScope.Value, this.Service.RequestedServerVersion);
            }
        }

        /// <summary>
        /// Writes the elements to XML.
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
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.AddDelegate;
        }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.AddDelegateResponse;
        }

        /// <summary>
        /// Creates the response.
        /// </summary>
        /// <returns>Service response.</returns>
        internal override DelegateManagementResponse CreateResponse()
        {
            return new DelegateManagementResponse(true, this.delegateUsers);
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
