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