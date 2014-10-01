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
