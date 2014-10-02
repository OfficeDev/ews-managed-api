#region License

// Exchange Web Services Managed API
// 
// Copyright (c) Microsoft Corporation
// All rights reserved. 
//
// MIT License
//
// Permission is hereby granted, free of charge, to any person obtaining a copy of this
// software and associated documentation files (the "Software"), to deal in the Software
// without restriction, including without limitation the rights to use, copy, modify, merge,
// publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
// to whom the Software is furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all copies or
// substantial portions of the Software.

// THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
// INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
// PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
// FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
// OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
// DEALINGS IN THE SOFTWARE.

#endregion

//-----------------------------------------------------------------------
// <summary>Defines the SendItemRequest class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a SendItem request.
    /// </summary>
    internal sealed class SendItemRequest : MultiResponseServiceRequest<ServiceResponse>, IJsonSerializable
    {
        private IEnumerable<Item> items;
        private FolderId savedCopyDestinationFolderId;

        /// <summary>
        /// Asserts the valid.
        /// </summary>
        internal override void Validate()
        {
            base.Validate();
            EwsUtilities.ValidateParam(this.Items, "Items");

            if (this.SavedCopyDestinationFolderId != null)
            {
                this.SavedCopyDestinationFolderId.Validate(this.Service.RequestedServerVersion);
            }
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
            return EwsUtilities.GetEnumeratedObjectCount(this.Items);
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.SendItem;
        }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.SendItemResponse;
        }

        /// <summary>
        /// Gets the name of the response message XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetResponseMessageXmlElementName()
        {
            return XmlElementNames.SendItemResponseMessage;
        }

        /// <summary>
        /// Writes the attributes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
        {
            base.WriteAttributesToXml(writer);

            writer.WriteAttributeValue(
                XmlAttributeNames.SaveItemToFolder,
                this.SavedCopyDestinationFolderId != null);
        }

        /// <summary>
        /// Writes the elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.ItemIds);

            foreach (Item item in this.Items)
            {
                item.Id.WriteToXml(writer, XmlElementNames.ItemId);
            }

            writer.WriteEndElement(); // ItemIds

            if (this.SavedCopyDestinationFolderId != null)
            {
                writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.SavedItemFolderId);
                this.SavedCopyDestinationFolderId.WriteToXml(writer);
                writer.WriteEndElement();
            }
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

            jsonRequest.Add(XmlAttributeNames.SaveItemToFolder, this.SavedCopyDestinationFolderId != null);
            if (this.SavedCopyDestinationFolderId != null)
            {
                JsonObject targetFolderId = new JsonObject();
                targetFolderId.Add(XmlElementNames.BaseFolderId, this.SavedCopyDestinationFolderId.InternalToJson(service));
                jsonRequest.Add(XmlElementNames.SavedItemFolderId, targetFolderId);
            }

            List<object> idList = new List<object>();
            foreach (Item item in this.Items)
            {
                idList.Add(item.Id.InternalToJson(service));
            }

            jsonRequest.Add(XmlElementNames.ItemIds, idList.ToArray());

            return jsonRequest;
        }

        /// <summary>
        /// Gets the request version.
        /// </summary>
        /// <returns>Earliest Exchange version in which this request is supported.</returns>
        internal override ExchangeVersion GetMinimumRequiredServerVersion()
        {
            return ExchangeVersion.Exchange2007_SP1;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="SendItemRequest"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="errorHandlingMode"> Indicates how errors should be handled.</param>
        internal SendItemRequest(ExchangeService service, ServiceErrorHandling errorHandlingMode)
            : base(service, errorHandlingMode)
        {
        }

        /// <summary>
        /// Gets or sets the items.
        /// </summary>
        /// <value>The items.</value>
        public IEnumerable<Item> Items
        {
            get { return this.items; }
            set { this.items = value; }
        }

        /// <summary>
        /// Gets or sets the saved copy destination folder id.
        /// </summary>
        /// <value>The saved copy destination folder id.</value>
        public FolderId SavedCopyDestinationFolderId
        {
            get { return this.savedCopyDestinationFolderId; }
            set { this.savedCopyDestinationFolderId = value; }
        }
    }
}
