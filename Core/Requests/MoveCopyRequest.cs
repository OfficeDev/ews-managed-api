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
