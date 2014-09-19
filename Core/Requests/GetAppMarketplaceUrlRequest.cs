// ---------------------------------------------------------------------------
// <copyright file="GetAppMarketplaceUrlRequest.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

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