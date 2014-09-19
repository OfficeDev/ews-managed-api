// ---------------------------------------------------------------------------
// <copyright file="ConversationResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ConversationResponseType class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// 
    /// </summary>
    public sealed class ConversationResponse : ComplexProperty
    {
        /// <summary>
        /// Property set used to fetch items in the conversation.
        /// </summary>
        private PropertySet propertySet;

        /// <summary>
        /// Initializes a new instance of the <see cref="ConversationResponse"/> class.
        /// </summary>
        /// <param name="propertySet">The property set.</param>
        internal ConversationResponse(PropertySet propertySet)
        {
            this.propertySet = propertySet;
        }

        /// <summary>
        /// Gets the conversation id.
        /// </summary>
        public ConversationId ConversationId { get; internal set; }

        /// <summary>
        /// Gets the sync state.
        /// </summary>
        public string SyncState { get; internal set; }

        /// <summary>
        /// Gets the conversation nodes.
        /// </summary>
        public ConversationNodeCollection ConversationNodes { get; internal set; }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            switch (reader.LocalName)
            {
                case XmlElementNames.ConversationId:
                    this.ConversationId = new ConversationId();
                    this.ConversationId.LoadFromXml(reader, XmlElementNames.ConversationId);
                    return true;

                case XmlElementNames.SyncState:
                    this.SyncState = reader.ReadElementValue();
                    return true;

                case XmlElementNames.ConversationNodes:
                    this.ConversationNodes = new ConversationNodeCollection(this.propertySet);
                    this.ConversationNodes.LoadFromXml(reader, XmlElementNames.ConversationNodes);
                    return true;

                default:
                    return false;
            }
        }

        /// <summary>
        /// Loads from json.
        /// </summary>
        /// <param name="jsonProperty">The json property.</param>
        /// <param name="service">The service.</param>
        internal override void LoadFromJson(JsonObject jsonProperty, ExchangeService service)
        {
            this.ConversationId = new ConversationId();
            this.ConversationId.LoadFromJson(jsonProperty.ReadAsJsonObject(XmlElementNames.ConversationId), service);

            if (jsonProperty.ContainsKey(XmlElementNames.SyncState))
            {
                this.SyncState = jsonProperty.ReadAsString(XmlElementNames.SyncState);
            }

            this.ConversationNodes = new ConversationNodeCollection(this.propertySet);
            ((IJsonCollectionDeserializer)this.ConversationNodes).CreateFromJsonCollection(
                                                                        jsonProperty.ReadAsArray(XmlElementNames.ConversationNodes),
                                                                        service);
        }
    }
}
