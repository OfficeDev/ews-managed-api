// ---------------------------------------------------------------------------
// <copyright file="SetTeamMailboxRequest.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the SetTeamMailboxRequest class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;

    /// <summary>
    /// Represents a SetTeamMailbox request.
    /// </summary>
    internal sealed class SetTeamMailboxRequest : SimpleServiceRequestBase
    {
        /// <summary>
        /// TeamMailbox email address
        /// </summary>
        private EmailAddress emailAddress;

        /// <summary>
        /// SharePoint site URL
        /// </summary>
        private Uri sharePointSiteUrl;

        /// <summary>
        /// TeamMailbox lifecycle state
        /// </summary>
        private TeamMailboxLifecycleState state;

        /// <summary>
        /// Initializes a new instance of the <see cref="SetTeamMailboxRequest"/> class.
        /// </summary>
        /// <param name="service">The service</param>
        /// <param name="emailAddress">TeamMailbox email address</param>
        /// <param name="sharePointSiteUrl">SharePoint site URL</param>
        /// <param name="state">TeamMailbox state</param>
        internal SetTeamMailboxRequest(ExchangeService service, EmailAddress emailAddress, Uri sharePointSiteUrl, TeamMailboxLifecycleState state)
            : base(service)
        {
            if (emailAddress == null)
            {
                throw new ArgumentNullException("emailAddress");
            }

            if (sharePointSiteUrl == null)
            {
                throw new ArgumentNullException("sharePointSiteUrl");
            }

            this.emailAddress = emailAddress;
            this.sharePointSiteUrl = sharePointSiteUrl;
            this.state = state;
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.SetTeamMailbox;
        }

        /// <summary>
        /// Writes XML elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            this.emailAddress.WriteToXml(writer, XmlNamespace.Messages, XmlElementNames.EmailAddress);
            writer.WriteElementValue(XmlNamespace.Messages, XmlElementNames.SharePointSiteUrl, this.sharePointSiteUrl.ToString());
            writer.WriteElementValue(XmlNamespace.Messages, XmlElementNames.State, this.state.ToString());
        }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.SetTeamMailboxResponse;
        }

        /// <summary>
        /// Parses the response.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>Response object.</returns>
        internal override object ParseResponse(EwsServiceXmlReader reader)
        {
            ServiceResponse response = new ServiceResponse();
            response.LoadFromXml(reader, GetResponseXmlElementName());
            return response;
        }

        /// <summary>
        /// Gets the request version.
        /// </summary>
        /// <returns>Earliest Exchange version in which this request is supported.</returns>
        internal override ExchangeVersion GetMinimumRequiredServerVersion()
        {
            return ExchangeVersion.Exchange2013;
        }

        /// <summary>
        /// Executes this request.
        /// </summary>
        /// <returns>Service response.</returns>
        internal ServiceResponse Execute()
        {
            ServiceResponse serviceResponse = (ServiceResponse)this.InternalExecute();
            serviceResponse.ThrowIfNecessary();
            return serviceResponse;
        }
    }
}