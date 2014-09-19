// ---------------------------------------------------------------------------
// <copyright file="UnpinTeamMailboxRequest.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the UnpinTeamMailboxRequest class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;

    /// <summary>
    /// Represents a UnpinTeamMailbox request.
    /// </summary>
    internal sealed class UnpinTeamMailboxRequest : SimpleServiceRequestBase
    {
        /// <summary>
        /// TeamMailbox email address
        /// </summary>
        private readonly EmailAddress emailAddress;

        /// <summary>
        /// Initializes a new instance of the <see cref="UnpinTeamMailboxRequest"/> class.
        /// </summary>
        /// <param name="service">The service</param>
        /// <param name="emailAddress">TeamMailbox email address</param>
        public UnpinTeamMailboxRequest(ExchangeService service, EmailAddress emailAddress)
            : base(service)
        {
            if (emailAddress == null)
            {
                throw new ArgumentNullException("emailAddress");
            }

            this.emailAddress = emailAddress;
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.UnpinTeamMailbox;
        }

        /// <summary>
        /// Writes XML elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            this.emailAddress.WriteToXml(writer, XmlNamespace.Messages, XmlElementNames.EmailAddress);
        }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.UnpinTeamMailboxResponse;
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