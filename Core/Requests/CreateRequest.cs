// ---------------------------------------------------------------------------
// <copyright file="CreateRequest.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the CreateRequest class.</summary>
//-----------------------------------------------------------------------

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
    internal abstract class CreateRequest<TServiceObject, TResponse> : MultiResponseServiceRequest<TResponse>, IJsonSerializable
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
        /// Creates a JSON representation of this object.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <returns>
        /// A Json value (either a JsonObject, an array of Json values, or a Json primitive)
        /// </returns>
        object IJsonSerializable.ToJson(ExchangeService service)
        {
            JsonObject jsonRequest = new JsonObject();

            if (this.ParentFolderId != null)
            {
                JsonObject targetFolderId = new JsonObject();
                targetFolderId.Add(XmlElementNames.BaseFolderId, this.ParentFolderId.InternalToJson(service));

                jsonRequest.Add(this.GetParentFolderXmlElementName(), targetFolderId);
            }

            List<object> jsonServiceObjects = new List<object>();
            foreach (ServiceObject obj in this.objects)
            {
                jsonServiceObjects.Add(obj.ToJson(service, false));
            }

            jsonRequest.Add(this.GetObjectCollectionXmlElementName(), jsonServiceObjects.ToArray());

            this.AddJsonProperties(jsonRequest, service);
            
            return jsonRequest;
        }

        /// <summary>
        /// Adds the json properties.
        /// </summary>
        /// <param name="jsonRequest">The json request.</param>
        /// <param name="service">The service.</param>
        internal virtual void AddJsonProperties(JsonObject jsonRequest, ExchangeService service)
        {
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
