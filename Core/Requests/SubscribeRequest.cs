// ---------------------------------------------------------------------------
// <copyright file="SubscribeRequest.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the SubscribeRequest class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System.Collections.Generic;
    using System.Linq;

    /// <summary>
    /// Represents an abstract Subscribe request.
    /// </summary>
    /// <typeparam name="TSubscription">The type of the subscription.</typeparam>
    internal abstract class SubscribeRequest<TSubscription> : MultiResponseServiceRequest<SubscribeResponse<TSubscription>>, IJsonSerializable
        where TSubscription : SubscriptionBase
    {
        /// <summary>
        /// Validate request.
        /// </summary>
        internal override void Validate()
        {
            base.Validate();
            EwsUtilities.ValidateParam(this.FolderIds, "FolderIds");
            EwsUtilities.ValidateParamCollection(this.EventTypes, "EventTypes");
            this.FolderIds.Validate(this.Service.RequestedServerVersion);

            // Check that caller isn't trying to subscribe to Status events.
            if (this.EventTypes.Count<EventType>(eventType => (eventType == EventType.Status)) > 0)
            {
                throw new ServiceValidationException(Strings.CannotSubscribeToStatusEvents);
            }

            // If Watermark was specified, make sure it's not a blank string.
            if (!string.IsNullOrEmpty(this.Watermark))
            {
                EwsUtilities.ValidateNonBlankStringParam(this.Watermark, "Watermark");
            }

            this.EventTypes.ForEach(eventType => EwsUtilities.ValidateEnumVersionValue(eventType, this.Service.RequestedServerVersion));
        }

        /// <summary>
        /// Gets the name of the subscription XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal abstract string GetSubscriptionXmlElementName();

        /// <summary>
        /// Gets the expected response message count.
        /// </summary>
        /// <returns>Number of expected response messages.</returns>
        internal override int GetExpectedResponseMessageCount()
        {
            return 1;
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.Subscribe;
        }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.SubscribeResponse;
        }

        /// <summary>
        /// Gets the name of the response message XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetResponseMessageXmlElementName()
        {
            return XmlElementNames.SubscribeResponseMessage;
        }

        /// <summary>
        /// Internal method to write XML elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal abstract void InternalWriteElementsToXml(EwsServiceXmlWriter writer);

        /// <summary>
        /// Writes XML elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteStartElement(XmlNamespace.Messages, this.GetSubscriptionXmlElementName());

            if (this.FolderIds.Count == 0)
            {
                writer.WriteAttributeValue(
                    XmlAttributeNames.SubscribeToAllFolders,
                    true);
            }

            this.FolderIds.WriteToXml(
                writer,
                XmlNamespace.Types,
                XmlElementNames.FolderIds);

            writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.EventTypes);
            foreach (EventType eventType in this.EventTypes)
            {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.EventType,
                    eventType);
            }
            writer.WriteEndElement();

            if (!string.IsNullOrEmpty(this.Watermark))
            {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.Watermark,
                    this.Watermark);
            }

            this.InternalWriteElementsToXml(writer);

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

            JsonObject jsonSubscribeRequest = new JsonObject();

            jsonSubscribeRequest.AddTypeParameter(this.GetSubscriptionXmlElementName());
            jsonSubscribeRequest.Add(XmlElementNames.EventTypes, this.EventTypes.ToArray());

            if (this.FolderIds.Count > 0)
            {
                jsonSubscribeRequest.Add(XmlElementNames.FolderIds, this.FolderIds.InternalToJson(service));
            }
            else
            {
                jsonSubscribeRequest.Add(XmlAttributeNames.SubscribeToAllFolders, true);
            }

            if (!string.IsNullOrEmpty(this.Watermark))
            {
                jsonSubscribeRequest.Add(XmlElementNames.Watermark, this.Watermark);
            }

            this.AddJsonProperties(jsonSubscribeRequest, service);

            jsonRequest.Add(XmlElementNames.SubscriptionRequest, jsonSubscribeRequest);

            return jsonRequest;
        }

        /// <summary>
        /// Adds the json properties.
        /// </summary>
        /// <param name="jsonSubscribeRequest">The json subscribe request.</param>
        /// <param name="service">The service.</param>
        internal abstract void AddJsonProperties(JsonObject jsonSubscribeRequest, ExchangeService service);

        /// <summary>
        /// Initializes a new instance of the <see cref="SubscribeRequest&lt;TSubscription&gt;"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        internal SubscribeRequest(ExchangeService service)
            : base(service, ServiceErrorHandling.ThrowOnError)
        {
            this.FolderIds = new FolderIdWrapperList();
            this.EventTypes = new List<EventType>();
        }

        /// <summary>
        /// Gets the folder ids.
        /// </summary>
        public FolderIdWrapperList FolderIds
        {
            get; private set;
        }

        /// <summary>
        /// Gets the event types.
        /// </summary>
        public List<EventType> EventTypes
        {
            get; private set;
        }

        /// <summary>
        /// Gets or sets the watermark.
        /// </summary>
        public string Watermark
        {
            get; set;
        }
    }
}
