// ---------------------------------------------------------------------------
// <copyright file="GetConversationItemsResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the GetConversationItemsResponse class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System.Collections.ObjectModel;

    /// <summary>
    /// Represents the response to a GetConversationItems operation.
    /// </summary>
    public sealed class GetConversationItemsResponse : ServiceResponse
    {
        private PropertySet propertySet;

        /// <summary>
        /// Initializes a new instance of the <see cref="GetConversationItemsResponse"/> class.
        /// </summary>
        /// <param name="propertySet">The property set.</param>
        internal GetConversationItemsResponse(PropertySet propertySet)
            : base()
        {
            this.propertySet = propertySet;
        }

        /// <summary>
        /// Gets or sets the conversation.
        /// </summary>
        /// <value>The conversation.</value>
        public ConversationResponse Conversation { get; set; }

        /// <summary>
        /// Read Conversations from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
        {
            this.Conversation = new ConversationResponse(this.propertySet);

            reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.Conversation);
            this.Conversation.LoadFromXml(reader, XmlNamespace.Messages, XmlElementNames.Conversation);
        }

        /// <summary>
        /// Reads response elements from Json.
        /// </summary>
        /// <param name="responseObject">The response object.</param>
        /// <param name="service">The service.</param>
        internal override void ReadElementsFromJson(JsonObject responseObject, ExchangeService service)
        {
            this.Conversation = new ConversationResponse(this.propertySet);
            this.Conversation.LoadFromJson(responseObject.ReadAsJsonObject(XmlElementNames.Conversation), service);
        }
    }
}
