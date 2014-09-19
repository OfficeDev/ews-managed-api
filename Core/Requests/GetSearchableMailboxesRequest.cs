// ---------------------------------------------------------------------------
// <copyright file="GetSearchableMailboxesRequest.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the GetSearchableMailboxesRequest class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a GetSearchableMailboxesRequest request.
    /// </summary>
    internal sealed class GetSearchableMailboxesRequest : SimpleServiceRequestBase, IJsonSerializable
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="GetSearchableMailboxesRequest"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        internal GetSearchableMailboxesRequest(ExchangeService service)
            : base(service)
        {
        }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.GetSearchableMailboxesResponse;
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.GetSearchableMailboxes;
        }

        /// <summary>
        /// Parses the response.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>Response object.</returns>
        internal override object ParseResponse(EwsServiceXmlReader reader)
        {
            GetSearchableMailboxesResponse response = new GetSearchableMailboxesResponse();
            response.LoadFromXml(reader, this.GetResponseXmlElementName());
            return response;
        }

        /// <summary>
        /// Writes XML elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteElementValue(XmlNamespace.Messages, XmlElementNames.SearchFilter, this.SearchFilter ?? string.Empty);
            writer.WriteElementValue(XmlNamespace.Messages, XmlElementNames.ExpandGroupMembership, this.ExpandGroupMembership.ToString().ToLower());
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
        internal GetSearchableMailboxesResponse Execute()
        {
            GetSearchableMailboxesResponse serviceResponse = (GetSearchableMailboxesResponse)this.InternalExecute();
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
        /// Search filter
        /// </summary>
        public string SearchFilter { get; set; }

        /// <summary>
        /// Expand group membership
        /// </summary>
        public bool ExpandGroupMembership { get; set; }
    }
}