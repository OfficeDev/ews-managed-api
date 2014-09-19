// ---------------------------------------------------------------------------
// <copyright file="ApplyConversationActionRequest.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ApplyConversationActionRequest class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a request to a Apply Conversation Action operation
    /// </summary>
    internal sealed class ApplyConversationActionRequest : MultiResponseServiceRequest<ServiceResponse>, IJsonSerializable
    {
        private List<ConversationAction> conversationActions = new List<ConversationAction>();

        public List<ConversationAction> ConversationActions
        {
            get { return this.conversationActions; }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ApplyConversationActionRequest"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="errorHandlingMode">Indicates how errors should be handled.</param>
        internal ApplyConversationActionRequest(ExchangeService service, ServiceErrorHandling errorHandlingMode)
            : base(service, errorHandlingMode)
        {
        }

        /// <summary>
        /// Creates the service response.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="responseIndex">Index of the response.</param>
        /// <returns>Service response.</returns>
        internal override ServiceResponse CreateServiceResponse(ExchangeService service, int responseIndex)
        {
            return new ServiceResponse();
        }

        /// <summary>
        /// Gets the expected response message count.
        /// </summary>
        /// <returns>Number of expected response messages.</returns>
        internal override int GetExpectedResponseMessageCount()
        {
            return this.conversationActions.Count;
        }

        /// <summary>
        /// Validate request.
        /// </summary>
        internal override void Validate()
        {
            base.Validate();
            EwsUtilities.ValidateParamCollection(this.conversationActions, "conversationActions");
            for (int iAction = 0; iAction < this.ConversationActions.Count; iAction++)
            {
                this.ConversationActions[iAction].Validate();
            }
        }

        /// <summary>
        /// Writes XML elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteStartElement(
                XmlNamespace.Messages,
                XmlElementNames.ConversationActions);
            for (int iAction = 0; iAction < this.ConversationActions.Count; iAction++)
            {
                this.ConversationActions[iAction].WriteElementsToXml(writer);
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
            List<object> actions = new List<object>();

            foreach (ConversationAction action in this.conversationActions)
            {
                actions.Add(((IJsonSerializable)action).ToJson(service));
            }

            jsonRequest.Add(XmlElementNames.ConversationActions, actions.ToArray());

            return jsonRequest;
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.ApplyConversationAction;
        }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.ApplyConversationActionResponse;
        }

        /// <summary>
        /// Gets the name of the response message XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetResponseMessageXmlElementName()
        {
            return XmlElementNames.ApplyConversationActionResponseMessage;
        }

        /// <summary>
        /// Gets the request version.
        /// </summary>
        /// <returns>Earliest Exchange version in which this request is supported.</returns>
        internal override ExchangeVersion GetMinimumRequiredServerVersion()
        {
            return ExchangeVersion.Exchange2010_SP1;
        }
    }
}
