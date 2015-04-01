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
    /// Represents a request to a GetConversationItems operation
    /// </summary>
    internal sealed class GetConversationItemsRequest : MultiResponseServiceRequest<GetConversationItemsResponse>, IJsonSerializable
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="GetConversationItemsRequest"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="errorHandlingMode">Error handling mode.</param>
        internal GetConversationItemsRequest(ExchangeService service, ServiceErrorHandling errorHandlingMode)
            : base(service, errorHandlingMode)
        {
        }

        /// <summary>
        /// Gets or sets the conversations.
        /// </summary>
        internal List<ConversationRequest> Conversations { get; set; }

        /// <summary>
        /// Gets or sets the item properties.
        /// </summary>
        internal PropertySet ItemProperties { get; set; }

        /// <summary>
        /// Gets or sets the folders to ignore.
        /// </summary>
        internal FolderIdCollection FoldersToIgnore { get; set; }

        /// <summary>
        /// Gets or sets the maximum number of items to return.
        /// </summary>
        internal int? MaxItemsToReturn { get; set; }

        internal ConversationSortOrder? SortOrder { get; set; }

        /// <summary>
        /// Gets or sets the mailbox search location to include in the search.
        /// </summary>
        internal MailboxSearchLocation? MailboxScope { get; set; }

        /// <summary>
        /// Validate request.
        /// </summary>
        internal override void Validate()
        {
            base.Validate();

            // SearchScope is only valid for Exchange2013 or higher
            //
            if (this.MailboxScope.HasValue &&
                this.Service.RequestedServerVersion < ExchangeVersion.Exchange2013)
            {
                throw new ServiceVersionException(
                    string.Format(
                        Strings.ParameterIncompatibleWithRequestVersion,
                        "MailboxScope",
                        ExchangeVersion.Exchange2013));
            }
        }

        /// <summary>
        /// Writes XML attributes.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
        {
            base.WriteAttributesToXml(writer);
        }

        /// <summary>
        /// Writes XML elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            this.ItemProperties.WriteToXml(writer, ServiceObjectType.Item);

            this.FoldersToIgnore.WriteToXml(writer, XmlNamespace.Messages, XmlElementNames.FoldersToIgnore);

            if (this.MaxItemsToReturn.HasValue)
            {
                writer.WriteElementValue(XmlNamespace.Messages, XmlElementNames.MaxItemsToReturn, this.MaxItemsToReturn.Value);
            }

            if (this.SortOrder.HasValue)
            {
                writer.WriteElementValue(XmlNamespace.Messages, XmlElementNames.SortOrder, this.SortOrder.Value);
            }

            if (this.MailboxScope.HasValue)
            {
                writer.WriteElementValue(XmlNamespace.Messages, XmlElementNames.MailboxScope, this.MailboxScope.Value);
            }

            writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.Conversations);
            this.Conversations.ForEach((conversation) => conversation.WriteToXml(writer, XmlElementNames.Conversation));
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

            this.ItemProperties.WriteGetShapeToJson(jsonRequest, service, ServiceObjectType.Item);

            if (this.FoldersToIgnore.Count > 0)
            {
                jsonRequest.Add(XmlElementNames.FoldersToIgnore, this.FoldersToIgnore.InternalToJson(service));
            }

            if (this.MaxItemsToReturn.HasValue)
            {
                jsonRequest.Add(XmlElementNames.MaxItemsToReturn, this.MaxItemsToReturn.Value);
            }

            if (this.SortOrder.HasValue)
            {
                jsonRequest.Add(XmlElementNames.SortOrder, this.SortOrder.Value);
            }

            if (this.MailboxScope.HasValue)
            {
                jsonRequest.Add(XmlElementNames.MailboxScope, this.MailboxScope.Value);
            }

            List<object> jsonPropertyCollection = new List<object>();

            this.Conversations.ForEach((conversation) => jsonPropertyCollection.Add(conversation.InternalToJson(service)));

            jsonRequest.Add(XmlElementNames.Conversations, jsonPropertyCollection.ToArray());

            return jsonRequest;
        }

        /// <summary>
        /// Creates the service response.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="responseIndex">Index of the response.</param>
        /// <returns>Service response.</returns>
        internal override GetConversationItemsResponse CreateServiceResponse(ExchangeService service, int responseIndex)
        {
            return new GetConversationItemsResponse(this.ItemProperties);
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.GetConversationItems;
        }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.GetConversationItemsResponse;
        }

        /// <summary>
        /// Gets the name of the response message XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetResponseMessageXmlElementName()
        {
            return XmlElementNames.GetConversationItemsResponseMessage;
        }

        /// <summary>
        /// Gets the request version.
        /// </summary>
        /// <returns>Earliest Exchange version in which this request is supported.</returns>
        internal override ExchangeVersion GetMinimumRequiredServerVersion()
        {
            return ExchangeVersion.Exchange2013;
        }

        /// <summary>
        /// Gets the expected response message count.
        /// </summary>
        /// <returns>Number of expected response messages.</returns>
        internal override int GetExpectedResponseMessageCount()
        {
            return this.Conversations.Count;
        }
    }
}