// ---------------------------------------------------------------------------
// <copyright file="ConversationRequest.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ConversationRequest class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// 
    /// </summary>
    public sealed class ConversationRequest : ComplexProperty, ISelfValidate, IJsonSerializable
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ConversationRequest"/> class.
        /// </summary>
        public ConversationRequest()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ConversationRequest"/> class.
        /// </summary>
        /// <param name="conversationId">The conversation id.</param>
        /// <param name="syncState">State of the sync.</param>
        public ConversationRequest(ConversationId conversationId, string syncState)
        {
            this.ConversationId = conversationId;
            this.SyncState = syncState;
        }

        /// <summary>
        /// Gets or sets the conversation id.
        /// </summary>
        public ConversationId ConversationId { get; set; }

        /// <summary>
        /// Gets or sets the sync state representing the current state of the conversation for synchronization purposes.
        /// </summary>
        public string SyncState { get; set; }

        /// <summary>
        /// Writes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="xmlElementName">Name of the XML element.</param>
        internal override void WriteToXml(EwsServiceXmlWriter writer, string xmlElementName)
        {
            writer.WriteStartElement(XmlNamespace.Types, xmlElementName);

            this.ConversationId.WriteToXml(writer);

            if (this.SyncState != null)
            {
                writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.SyncState, this.SyncState);
            }

            writer.WriteEndElement();
        }

        /// <summary>
        /// Serializes the property to a Json value.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <returns>
        /// A Json value (either a JsonObject, an array of Json values, or a Json primitive)
        /// </returns>
        internal override object InternalToJson(ExchangeService service)
        {
            JsonObject jsonProperty = new JsonObject();
            
            jsonProperty.Add(XmlElementNames.ConversationId, this.ConversationId.InternalToJson(service));
            if (!string.IsNullOrEmpty(this.SyncState))
            {
                jsonProperty.Add(XmlElementNames.SyncState, this.SyncState);
            }

            return jsonProperty;
        }

        /// <summary>
        /// Validates this instance.
        /// </summary>
        internal override void InternalValidate()
        {
            EwsUtilities.ValidateParam(this.ConversationId, "ConversationId");
        }
    }
}
