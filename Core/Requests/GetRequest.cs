// ---------------------------------------------------------------------------
// <copyright file="GetRequest.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the GetRequest class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents an abstract Get request.
    /// </summary>
    /// <typeparam name="TServiceObject">The type of the service object.</typeparam>
    /// <typeparam name="TResponse">The type of the response.</typeparam>
    internal abstract class GetRequest<TServiceObject, TResponse> : MultiResponseServiceRequest<TResponse>, IJsonSerializable
        where TServiceObject : ServiceObject
        where TResponse : ServiceResponse
    {
        private PropertySet propertySet;

        /// <summary>
        /// Initializes a new instance of the <see cref="GetRequest&lt;TServiceObject, TResponse&gt;"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="errorHandlingMode"> Indicates how errors should be handled.</param>
        internal GetRequest(ExchangeService service, ServiceErrorHandling errorHandlingMode)
            : base(service, errorHandlingMode)
        {
        }

        /// <summary>
        /// Validate request.
        /// </summary>
        internal override void Validate()
        {
            base.Validate();
            EwsUtilities.ValidateParam(this.PropertySet, "PropertySet");
            this.PropertySet.ValidateForRequest(this, false /*summaryPropertiesOnly*/);
        }

        /// <summary>
        /// Gets the type of the service object this request applies to.
        /// </summary>
        /// <returns>The type of service object the request applies to.</returns>
        internal abstract ServiceObjectType GetServiceObjectType();

        /// <summary>
        /// Writes XML elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            this.propertySet.WriteToXml(writer, this.GetServiceObjectType());
        }

        /// <summary>
        /// Gets or sets the property set.
        /// </summary>
        /// <value>The property set.</value>
        public PropertySet PropertySet
        {
            get { return this.propertySet; }
            set { this.propertySet = value; }
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

            this.propertySet.WriteGetShapeToJson(jsonRequest, service, this.GetServiceObjectType());
            this.AddIdsToRequest(jsonRequest, service);

            return jsonRequest;
        }

        internal abstract void AddIdsToRequest(JsonObject jsonRequest, ExchangeService service);
    }
}
