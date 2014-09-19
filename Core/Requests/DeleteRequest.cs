// ---------------------------------------------------------------------------
// <copyright file="DeleteRequest.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the DeleteRequest class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents an abstract Delete request.
    /// </summary>
    /// <typeparam name="TResponse">The type of the response.</typeparam>
    internal abstract class DeleteRequest<TResponse> : MultiResponseServiceRequest<TResponse>, IJsonSerializable
        where TResponse : ServiceResponse
    {
        /// <summary>
        /// Delete mode. Default is SoftDelete.
        /// </summary>
        private DeleteMode deleteMode = DeleteMode.SoftDelete;

        /// <summary>
        /// Initializes a new instance of the <see cref="DeleteRequest&lt;TResponse&gt;"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="errorHandlingMode"> Indicates how errors should be handled.</param>
        internal DeleteRequest(ExchangeService service, ServiceErrorHandling errorHandlingMode)
            : base(service, errorHandlingMode)
        {
        }

        /// <summary>
        /// Writes XML attributes.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
        {
            base.WriteAttributesToXml(writer);

            writer.WriteAttributeValue(XmlAttributeNames.DeleteType, this.DeleteMode);
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
            JsonObject body = new JsonObject();

            body.Add(XmlAttributeNames.DeleteType, this.DeleteMode.ToString());

            this.InternalToJson(body);

            return body;
        }

        protected abstract void InternalToJson(JsonObject body);

        /// <summary>
        /// Gets or sets the delete mode.
        /// </summary>
        /// <value>The delete mode.</value>
        public DeleteMode DeleteMode
        {
            get { return this.deleteMode; }
            set { this.deleteMode = value; }
        }
    }
}
