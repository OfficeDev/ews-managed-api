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
// <summary>Defines the GetAppMarketplaceUrlRequest class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Text;

    /// <summary>
    /// Represents a GetAppMarketplaceUrl request.
    /// </summary>
    internal sealed class GetAppMarketplaceUrlRequest : SimpleServiceRequestBase
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="GetAppMarketplaceUrlRequest"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        internal GetAppMarketplaceUrlRequest(ExchangeService service)
            : base(service)
        {
        }

        /// <summary>
        /// Gets or sets the api version supported by the client.
        /// This is used by EWS to generate a market place url with the correct version filter.
        /// </summary>
        /// <value>The Api version supported.</value>
        internal string ApiVersionSupported
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets the Schema version supported by the client.
        /// This is used by EWS to generate a market place url with the correct version filter.
        /// </summary>
        /// <value>The schema version supported.</value>
        internal string SchemaVersionSupported
        {
            get;
            set;
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.GetAppMarketplaceUrlRequest;
        }

        /// <summary>
        /// Validate request.
        /// </summary>
        internal override void Validate()
        {
            base.Validate();
            EwsUtilities.ValidateNonBlankStringParamAllowNull(this.ApiVersionSupported, "ApiVersionSupported");
            EwsUtilities.ValidateNonBlankStringParamAllowNull(this.SchemaVersionSupported, "SchemaVersionSupported");
        }

        /// <summary>
        /// Writes XML elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            if (!string.IsNullOrEmpty(this.ApiVersionSupported))
            {
                writer.WriteElementValue(XmlNamespace.Messages, "ApiVersionSupported", this.ApiVersionSupported);
            }

            if (!string.IsNullOrEmpty(this.SchemaVersionSupported))
            {
                writer.WriteElementValue(XmlNamespace.Messages, "SchemaVersionSupported", this.SchemaVersionSupported);
            }
        }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.GetAppMarketplaceUrlResponse;
        }

        /// <summary>
        /// Parses the response.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>Response object.</returns>
        internal override object ParseResponse(EwsServiceXmlReader reader)
        {
            GetAppMarketplaceUrlResponse response = new GetAppMarketplaceUrlResponse();
            response.LoadFromXml(reader, XmlElementNames.GetAppMarketplaceUrlResponse);
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
        internal GetAppMarketplaceUrlResponse Execute()
        {
            GetAppMarketplaceUrlResponse serviceResponse = (GetAppMarketplaceUrlResponse)this.InternalExecute();
            serviceResponse.ThrowIfNecessary();
            return serviceResponse;
        }
    }
}