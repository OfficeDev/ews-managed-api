// ---------------------------------------------------------------------------
// <copyright file="GetUserUnifiedGroupsRequest.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the GetUserUnifiedGroupsRequest class.</summary>
//-----------------------------------------------------------------------
namespace Microsoft.Exchange.WebServices.Data.Groups
{
    using System;
    using System.Collections.Generic;
    using System.Text;
    using Microsoft.Exchange.WebServices.Data.Groups;

    /// <summary>
    /// Represents a request to a GetUserUnifiedGroupsRequest operation
    /// </summary>
    internal sealed class GetUserUnifiedGroupsRequest : SimpleServiceRequestBase, IJsonSerializable
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="GetUserUnifiedGroupsRequest"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        internal GetUserUnifiedGroupsRequest(ExchangeService service) : base(service)
        {
        }

        /// <summary>
        /// Gets or sets the RequestedUnifiedGroupsSet
        /// </summary>
        public IEnumerable<RequestedUnifiedGroupsSet> RequestedUnifiedGroupsSets { get; set; }

        /// <summary>
        /// Gets or sets the UserSmptAddress
        /// </summary>
        public string UserSmtpAddress { get; set; }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.GetUserUnifiedGroupsResponseMessage;
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.GetUserUnifiedGroups;
        }

        /// <summary>
        /// Parses the response.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>Response object.</returns>
        internal override object ParseResponse(EwsServiceXmlReader reader)
        {
            GetUserUnifiedGroupsResponse response = new GetUserUnifiedGroupsResponse();
            response.LoadFromXml(reader, GetResponseXmlElementName());
            return response;
        }

        /// <summary>
        /// Writes XML elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.RequestedGroupsSets);

            if (this.RequestedUnifiedGroupsSets != null)
            { 
                this.RequestedUnifiedGroupsSets.ForEach((unifiedGroupsSet) => unifiedGroupsSet.WriteToXml(writer, XmlElementNames.RequestedUnifiedGroupsSetItem));
            }

            writer.WriteEndElement();

            if (!string.IsNullOrEmpty(this.UserSmtpAddress))
            {
                writer.WriteElementValue(XmlNamespace.NotSpecified, XmlElementNames.UserSmtpAddress, this.UserSmtpAddress);
            }
        }

        /// <summary>
        /// Gets the request version.
        /// </summary>
        /// <returns>Earliest Exchange version in which this request is supported.</returns>
        internal override ExchangeVersion GetMinimumRequiredServerVersion()
        {
            return ExchangeVersion.Exchange2013_SP1;
        }

        /// <summary>
        /// Executes this request.
        /// </summary>
        /// <returns>Service response.</returns>
        internal GetUserUnifiedGroupsResponse Execute()
        {
            return (GetUserUnifiedGroupsResponse)this.InternalExecute();
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
            JsonObject jsonRequest = new JsonObject();

            List<object> jsonPropertyCollection = new List<object>();
            if (this.RequestedUnifiedGroupsSets != null)
            {
                this.RequestedUnifiedGroupsSets.ForEach((unifiedGroupsSet) => jsonPropertyCollection.Add(unifiedGroupsSet.InternalToJson(service)));
                jsonRequest.Add(XmlElementNames.RequestedGroupsSets, jsonPropertyCollection.ToArray());
            }

            if (!string.IsNullOrEmpty(this.UserSmtpAddress))
            {
                jsonRequest.Add(XmlElementNames.SmtpAddress, this.UserSmtpAddress);
            }

            return jsonRequest;
        }
    }
}