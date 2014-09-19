// ---------------------------------------------------------------------------
// <copyright file="GetDiscoverySearchConfigurationRequest.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the GetDiscoverySearchConfigurationRequest class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a GetDiscoverySearchConfigurationRequest.
    /// </summary>
    internal sealed class GetDiscoverySearchConfigurationRequest : SimpleServiceRequestBase, IJsonSerializable
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="GetDiscoverySearchConfigurationRequest"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        internal GetDiscoverySearchConfigurationRequest(ExchangeService service)
            : base(service)
        {
        }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.GetDiscoverySearchConfigurationResponse;
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.GetDiscoverySearchConfiguration;
        }

        /// <summary>
        /// Parses the response.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>Response object.</returns>
        internal override object ParseResponse(EwsServiceXmlReader reader)
        {
            GetDiscoverySearchConfigurationResponse response = new GetDiscoverySearchConfigurationResponse();
            response.LoadFromXml(reader, this.GetResponseXmlElementName());
            return response;
        }

        /// <summary>
        /// Writes XML elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteElementValue(XmlNamespace.Messages, XmlElementNames.SearchId, this.SearchId ?? string.Empty);
            writer.WriteElementValue(XmlNamespace.Messages, XmlElementNames.ExpandGroupMembership, this.ExpandGroupMembership.ToString().ToLower());
            writer.WriteElementValue(XmlNamespace.Messages, XmlElementNames.InPlaceHoldConfigurationOnly, this.InPlaceHoldConfigurationOnly.ToString().ToLower());
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
        internal GetDiscoverySearchConfigurationResponse Execute()
        {
            GetDiscoverySearchConfigurationResponse serviceResponse = (GetDiscoverySearchConfigurationResponse)this.InternalExecute();
            return serviceResponse;
        }

        /// <summary>
        /// Creates a JSON representation of this object.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <returns>
        /// A Json value (either a JsonObject, an array of Json values, or a Json primitive)
        /// </returns>
        object IJsonSerializable.ToJson(ExchangeService service)
        {
            JsonObject jsonObject = new JsonObject();

            return jsonObject;
        }

        /// <summary>
        /// Search Id
        /// </summary>
        public string SearchId { get; set; }

        /// <summary>
        /// Expand group membership
        /// </summary>
        public bool ExpandGroupMembership { get; set; }

        /// <summary>
        /// In-Place hold configuration only
        /// </summary>
        public bool InPlaceHoldConfigurationOnly { get; set; }
    }
}