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