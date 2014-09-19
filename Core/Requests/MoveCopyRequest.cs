// ---------------------------------------------------------------------------
// <copyright file="MoveCopyRequest.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the MoveCopyRequest class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents an abstract Move/Copy request.
    /// </summary>
    /// <typeparam name="TServiceObject">The type of the service object.</typeparam>
    /// <typeparam name="TResponse">The type of the response.</typeparam>
    internal abstract class MoveCopyRequest<TServiceObject, TResponse> : MultiResponseServiceRequest<TResponse>, IJsonSerializable
        where TServiceObject : ServiceObject
        where TResponse : ServiceResponse
    {
        private FolderId destinationFolderId;

        /// <summary>
        /// Validates request.
        /// </summary>
        internal override void Validate()
        {
            EwsUtilities.ValidateParam(this.DestinationFolderId, "DestinationFolderId");
            this.DestinationFolderId.Validate(this.Service.RequestedServerVersion);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="MoveCopyRequest&lt;TServiceObject, TResponse&gt;"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="errorHandlingMode"> Indicates how errors should be handled.</param>
        internal MoveCopyRequest(ExchangeService service, ServiceErrorHandling errorHandlingMode)
            : base(service, errorHandlingMode)
        {
        }

        /// <summary>
        /// Writes the ids as XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal abstract void WriteIdsToXml(EwsServiceXmlWriter writer);

        /// <summary>
        /// Writes XML elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.ToFolderId);
            this.DestinationFolderId.WriteToXml(writer);
            writer.WriteEndElement();

            this.WriteIdsToXml(writer);
        }

        /// <summary>
        /// Gets or sets the destination folder id.
        /// </summary>
        /// <value>The destination folder id.</value>
        public FolderId DestinationFolderId
        {
            get { return this.destinationFolderId; }
            set { this.destinationFolderId = value; }
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

            JsonObject jsonToFolderId = new JsonObject();
            jsonToFolderId.Add(XmlElementNames.BaseFolderId, this.DestinationFolderId.InternalToJson(service));

            jsonObject.Add(XmlElementNames.ToFolderId, jsonToFolderId);

            this.AddIdsToJson(jsonObject, service);

            return jsonObject;
        }

        /// <summary>
        /// Adds the ids to json.
        /// </summary>
        /// <param name="jsonObject">The json object.</param>
        /// <param name="service">The service.</param>
        internal abstract void AddIdsToJson(JsonObject jsonObject, ExchangeService service);
    }
}
