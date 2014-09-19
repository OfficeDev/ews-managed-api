// ---------------------------------------------------------------------------
// <copyright file="GetPhoneCallRequest.cs" company="Microsoft">
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
    /// Represents a GetPhoneCall request.
    /// </summary>
    internal sealed class GetPhoneCallRequest : SimpleServiceRequestBase
    {
        private PhoneCallId id;

        /// <summary>
        /// Initializes a new instance of the <see cref="GetPhoneCallRequest"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        internal GetPhoneCallRequest(ExchangeService service)
            : base(service)
        {
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.GetPhoneCall;
        }

        /// <summary>
        /// Writes XML elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            this.id.WriteToXml(writer, XmlNamespace.Messages, XmlElementNames.PhoneCallId);
        }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.GetPhoneCallResponse;
        }

        /// <summary>
        /// Parses the response.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>Response object.</returns>
        internal override object ParseResponse(EwsServiceXmlReader reader)
        {
            GetPhoneCallResponse response = new GetPhoneCallResponse(Service);
            response.LoadFromXml(reader, XmlElementNames.GetPhoneCallResponse);
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
        internal GetPhoneCallResponse Execute()
        {
            GetPhoneCallResponse serviceResponse = (GetPhoneCallResponse)this.InternalExecute();
            serviceResponse.ThrowIfNecessary();
            return serviceResponse;
        }

        /// <summary>
        /// Gets or sets the Id of the phone call.
        /// </summary>
        internal PhoneCallId Id
        {
            get
            {
                return this.id;
            }

            set
            {
                this.id = value;
            }
        }
    }
}