// ---------------------------------------------------------------------------
// <copyright file="PlayOnPhoneRequest.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the PlayOnPhoneRequest class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a PlayOnPhone request.
    /// </summary>
    internal sealed class PlayOnPhoneRequest : SimpleServiceRequestBase, IJsonSerializable
    {
        private ItemId itemId;
        private string dialString;

        /// <summary>
        /// Initializes a new instance of the <see cref="PlayOnPhoneRequest"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        internal PlayOnPhoneRequest(ExchangeService service)
            : base(service)
        {
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.PlayOnPhone;
        }

        /// <summary>
        /// Writes XML elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            this.itemId.WriteToXml(writer, XmlNamespace.Messages, XmlElementNames.ItemId);
            writer.WriteElementValue(XmlNamespace.Messages, XmlElementNames.DialString, dialString);
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

            jsonRequest.Add(XmlElementNames.ItemId, this.ItemId.InternalToJson(service));
            jsonRequest.Add(XmlElementNames.DialString, this.dialString);

            return jsonRequest;
        }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.PlayOnPhoneResponse;
        }

        /// <summary>
        /// Parses the response.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>Response object.</returns>
        internal override object ParseResponse(EwsServiceXmlReader reader)
        {
            PlayOnPhoneResponse serviceResponse = new PlayOnPhoneResponse(this.Service);
            serviceResponse.LoadFromXml(reader, XmlElementNames.PlayOnPhoneResponse);
            return serviceResponse;
        }

        /// <summary>
        /// Parses the response.
        /// </summary>
        /// <param name="jsonBody">The json body.</param>
        /// <returns>Response object.</returns>
        internal override object ParseResponse(JsonObject jsonBody)
        {
            PlayOnPhoneResponse serviceResponse = new PlayOnPhoneResponse(this.Service);
            serviceResponse.LoadFromJson(jsonBody, this.Service);
            return serviceResponse;
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
        internal PlayOnPhoneResponse Execute()
        {
            PlayOnPhoneResponse serviceResponse = (PlayOnPhoneResponse)this.InternalExecute();
            serviceResponse.ThrowIfNecessary();
            return serviceResponse;
        }

        /// <summary>
        /// Gets or sets the item id of the message to play.
        /// </summary>
        internal ItemId ItemId
        {
            get
            {
                return this.itemId;
            }

            set
            {
                this.itemId = value;
            }
        }

        /// <summary>
        /// Gets or sets the dial string.
        /// </summary>
        internal string DialString
        {
            get
            {
                return this.dialString;
            }

            set
            {
                this.dialString = value;
            }
        }
    }
}