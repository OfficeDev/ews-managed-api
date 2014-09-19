// ---------------------------------------------------------------------------
// <copyright file="GetClientAccessTokenRequest.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a GetClientAccessToken request.
    /// </summary>
    internal sealed class GetClientAccessTokenRequest : MultiResponseServiceRequest<GetClientAccessTokenResponse>, IJsonSerializable
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="GetClientAccessTokenRequest"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="errorHandlingMode"> Indicates how errors should be handled.</param>
        internal GetClientAccessTokenRequest(ExchangeService service, ServiceErrorHandling errorHandlingMode)
            : base(service, errorHandlingMode)
        {
        }

        /// <summary>
        /// Creates the service response.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="responseIndex">Index of the response.</param>
        /// <returns>Response object.</returns>
        internal override GetClientAccessTokenResponse CreateServiceResponse(ExchangeService service, int responseIndex)
        {
            return new GetClientAccessTokenResponse(
                this.TokenRequests[responseIndex].Id,
                this.TokenRequests[responseIndex].TokenType);
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.GetClientAccessToken;
        }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>Xml element name.</returns>
        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.GetClientAccessTokenResponse;
        }

        /// <summary>
        /// Gets the name of the response message XML element.
        /// </summary>
        /// <returns>Xml element name.</returns>
        internal override string GetResponseMessageXmlElementName()
        {
            return XmlElementNames.GetClientAccessTokenResponseMessage;
        }

        /// <summary>
        /// Gets the expected response message count.
        /// </summary>
        /// <returns>Number of items in response.</returns>
        internal override int GetExpectedResponseMessageCount()
        {
            return TokenRequests.Length;
        }

        /// <summary>
        /// Writes XML elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.TokenRequests);

            foreach (ClientAccessTokenRequest tokenRequestInfo in this.TokenRequests)
            {
                writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.TokenRequest);
                writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Id, tokenRequestInfo.Id);
                writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.TokenType, tokenRequestInfo.TokenType);
                if (!string.IsNullOrEmpty(tokenRequestInfo.Scope))
                {
                    writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.HighlightTermScope, tokenRequestInfo.Scope);
                }

                writer.WriteEndElement();
            }

            writer.WriteEndElement();
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
            List<object> jsonArray = new List<object>();

            foreach (ClientAccessTokenRequest tokenRequestInfo in this.TokenRequests)
            {
                JsonObject innerRequest = new JsonObject();
                innerRequest.AddTypeParameter(XmlElementNames.TokenRequest);
                innerRequest.Add(XmlElementNames.Id, tokenRequestInfo.Id);
                innerRequest.Add(XmlElementNames.TokenType, tokenRequestInfo.TokenType);
                if (!string.IsNullOrEmpty(tokenRequestInfo.Scope))
                {
                    innerRequest.Add(XmlElementNames.HighlightTermScope, tokenRequestInfo.Scope);
                }

                jsonArray.Add(innerRequest);
            }

            JsonObject jsonRequest = new JsonObject();
            jsonRequest.Add(XmlElementNames.TokenRequests, jsonArray.ToArray());
            return jsonRequest;
        }

        /// <summary>
        /// Validate request.
        /// </summary>
        internal override void Validate()
        {
            base.Validate();

            if (this.TokenRequests == null || this.TokenRequests.Length == 0)
            {
                throw new ServiceValidationException(Strings.HoldIdParameterIsNotSpecified);
            }
        }

        /// <summary>
        /// Gets the request version.
        /// </summary>
        /// <returns>Earliest Exchange version in which this request is supported.</returns>
        internal override ExchangeVersion GetMinimumRequiredServerVersion()
        {
            return ExchangeVersion.Exchange2013;
        }

        internal ClientAccessTokenRequest[] TokenRequests
        {
            get;
            set;
        }
    }
}