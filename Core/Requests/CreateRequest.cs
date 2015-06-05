/*
 * Exchange Web Services Managed API
 *
 * Copyright (c) Microsoft Corporation
 * All rights reserved.
 *
 * MIT License
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this
 * software and associated documentation files (the "Software"), to deal in the Software
 * without restriction, including without limitation the rights to use, copy, modify, merge,
 * publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
 * to whom the Software is furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all copies or
 * substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
 * INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
 * PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
 * FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
 * OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
 * DEALINGS IN THE SOFTWARE.
 */

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents an abstract Create request.
    /// </summary>
    /// <typeparam name="TServiceObject">The type of the service object.</typeparam>
    /// <typeparam name="TResponse">The type of the response.</typeparam>
    internal abstract class CreateRequest<TServiceObject, TResponse> : MultiResponseServiceRequest<TResponse>
        where TServiceObject : ServiceObject
        where TResponse : ServiceResponse
    {
        private FolderId parentFolderId;
        private IEnumerable<TServiceObject> objects;

        /// <summary>
        /// Initializes a new instance of the <see cref="CreateRequest&lt;TServiceObject, TResponse&gt;"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="errorHandlingMode"> Indicates how errors should be handled.</param>
        protected CreateRequest(ExchangeService service, ServiceErrorHandling errorHandlingMode)
            : base(service, errorHandlingMode)
        {
        }

        /// <summary>
        /// Validate request.
        /// </summary>
        internal override void Validate()
        {
            base.Validate();
            if (this.ParentFolderId != null)
            {
                this.ParentFolderId.Validate(this.Service.RequestedServerVersion);
            }
        }

        /// <summary>
        /// Gets the expected response message count.
        /// </summary>
        /// <returns>Number of responses expected.</returns>
        internal override int GetExpectedResponseMessageCount()
        {
            return EwsUtilities.GetEnumeratedObjectCount(this.objects);
        }

        /// <summary>
        /// Gets the name of the parent folder XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal abstract string GetParentFolderXmlElementName();

        /// <summary>
        /// Gets the name of the object collection XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal abstract string GetObjectCollectionXmlElementName();

        /// <summary>
        /// Writes XML elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            if (this.ParentFolderId != null)
            {
                writer.WriteStartElement(XmlNamespace.Messages, this.GetParentFolderXmlElementName());
                this.ParentFolderId.WriteToXml(writer);
                writer.WriteEndElement();
            }

            writer.WriteStartElement(XmlNamespace.Messages, this.GetObjectCollectionXmlElementName());
            foreach (ServiceObject obj in this.objects)
            {
                obj.WriteToXml(writer);
            }
            writer.WriteEndElement();
        }

        /// <summary>
        /// Gets or sets the service objects.
        /// </summary>
        /// <value>The objects.</value>
        internal IEnumerable<TServiceObject> Objects
        {
            get { return this.objects; }
            set { this.objects = value; }
        }

        /// <summary>
        /// Gets or sets the parent folder id.
        /// </summary>
        /// <value>The parent folder id.</value>
        public FolderId ParentFolderId
        {
            get { return this.parentFolderId; }
            set { this.parentFolderId = value; }
        }
    }
}