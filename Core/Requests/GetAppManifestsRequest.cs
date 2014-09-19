// ---------------------------------------------------------------------------
// <copyright file="GetAppManifestsRequest.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the GetAppManifestsRequest class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a GetAppManifests request.
    /// </summary>
    internal sealed class GetAppManifestsRequest : SimpleServiceRequestBase
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="GetAppManifestsRequest"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        internal GetAppManifestsRequest(ExchangeService service)
            : base(service)
        {
        }

        /// <summary>
        /// Gets or sets the api version supported by the client.
        /// This tells Exchange service which app manifests should be returned based on the api version.
        /// </summary>
        /// <value>The Api version supported.</value>
        internal string ApiVersionSupported
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets the Schema version supported by the client.
        /// This tells Exchange service which app manifests should be returned based on the schema version.
        /// </summary>
        /// <value>The schema version supported.</value>
        internal string SchemaVersionSupported
        {
            get;
            set;
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
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.GetAppManifestsRequest;
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
            return XmlElementNames.GetAppManifestsResponse;
        }

        /// <summary>
        /// Parses the response.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>Response object.</returns>
        internal override object ParseResponse(EwsServiceXmlReader reader)
        {
            GetAppManifestsResponse response = new GetAppManifestsResponse();
            response.LoadFromXml(reader, XmlElementNames.GetAppManifestsResponse);
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
        internal GetAppManifestsResponse Execute()
        {
            GetAppManifestsResponse serviceResponse = (GetAppManifestsResponse)this.InternalExecute();
            serviceResponse.ThrowIfNecessary();
            return serviceResponse;
        }
    }
}