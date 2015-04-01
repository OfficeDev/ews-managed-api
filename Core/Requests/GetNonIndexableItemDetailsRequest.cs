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
    /// Represents a GetNonIndexableItemDetailsRequest request.
    /// </summary>
    internal sealed class GetNonIndexableItemDetailsRequest : SimpleServiceRequestBase, IJsonSerializable
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="GetNonIndexableItemDetailsRequest"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        internal GetNonIndexableItemDetailsRequest(ExchangeService service)
            : base(service)
        {
        }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.GetNonIndexableItemDetailsResponse;
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.GetNonIndexableItemDetails;
        }

        /// <summary>
        /// Validate request.
        /// </summary>
        internal override void Validate()
        {
            base.Validate();

            if (this.Mailboxes == null || this.Mailboxes.Length == 0)
            {
                throw new ServiceValidationException(Strings.MailboxesParameterIsNotSpecified);
            }
        }

        /// <summary>
        /// Parses the response.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>Response object.</returns>
        internal override object ParseResponse(EwsServiceXmlReader reader)
        {
            GetNonIndexableItemDetailsResponse response = new GetNonIndexableItemDetailsResponse();
            response.LoadFromXml(reader, GetResponseXmlElementName());
            return response;
        }

        /// <summary>
        /// Writes XML elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.Mailboxes);
            foreach (string mailbox in this.Mailboxes)
            {
                writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.LegacyDN, mailbox);
            }

            writer.WriteEndElement();

            if (this.PageSize != null && this.PageSize.HasValue)
            {
                writer.WriteElementValue(XmlNamespace.Messages, XmlElementNames.PageSize, this.PageSize.Value.ToString());
            }

            if (!string.IsNullOrEmpty(this.PageItemReference))
            {
                writer.WriteElementValue(XmlNamespace.Messages, XmlElementNames.PageItemReference, this.PageItemReference);
            }

            if (this.PageDirection != null && this.PageDirection.HasValue)
            {
                writer.WriteElementValue(XmlNamespace.Messages, XmlElementNames.PageDirection, this.PageDirection.Value.ToString());
            }

            writer.WriteElementValue(XmlNamespace.Messages, XmlElementNames.SearchArchiveOnly, this.SearchArchiveOnly);
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
        /// Executes this request.
        /// </summary>
        /// <returns>Service response.</returns>
        internal GetNonIndexableItemDetailsResponse Execute()
        {
            GetNonIndexableItemDetailsResponse serviceResponse = (GetNonIndexableItemDetailsResponse)this.InternalExecute();
            return serviceResponse;
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
            JsonObject jsonObject = new JsonObject();

            return jsonObject;
        }

        /// <summary>
        /// Mailboxes
        /// </summary>
        public string[] Mailboxes { get; set; }

        /// <summary>
        /// Page size
        /// </summary>
        public int? PageSize { get; set; }

        /// <summary>
        /// Page item reference
        /// </summary>
        public string PageItemReference { get; set; }

        /// <summary>
        /// Page direction
        /// </summary>
        public SearchPageDirection? PageDirection { get; set; }

        /// <summary>
        /// Whether to search archive only
        /// </summary>
        public bool SearchArchiveOnly { get; set; }
    }
}