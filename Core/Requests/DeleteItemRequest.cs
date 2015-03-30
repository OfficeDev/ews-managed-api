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
    /// Represents a DeleteItem request.
    /// </summary>
    internal sealed class DeleteItemRequest : DeleteRequest<ServiceResponse>
    {
        private ItemIdWrapperList itemIds = new ItemIdWrapperList();
        private AffectedTaskOccurrence? affectedTaskOccurrences;
        private SendCancellationsMode? sendCancellationsMode;

        /// <summary>
        /// Initializes a new instance of the <see cref="DeleteItemRequest"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="errorHandlingMode"> Indicates how errors should be handled.</param>
        internal DeleteItemRequest(ExchangeService service, ServiceErrorHandling errorHandlingMode)
            : base(service, errorHandlingMode)
        {
        }

        /// <summary>
        /// Validate request.
        /// </summary>
        internal override void Validate()
        {
            base.Validate();
            EwsUtilities.ValidateParam(this.ItemIds, "ItemIds");

            if (this.SuppressReadReceipts && this.Service.RequestedServerVersion < ExchangeVersion.Exchange2013)
            {
                throw new ServiceVersionException(
                    string.Format(
                        Strings.ParameterIncompatibleWithRequestVersion,
                        "SuppressReadReceipts",
                        ExchangeVersion.Exchange2013));
            }
        }

        /// <summary>
        /// Gets the expected response message count.
        /// </summary>
        /// <returns>Number of expected response messages.</returns>
        internal override int GetExpectedResponseMessageCount()
        {
            return this.itemIds.Count;
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
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.DeleteItem;
        }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.DeleteItemResponse;
        }

        /// <summary>
        /// Gets the name of the response message XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetResponseMessageXmlElementName()
        {
            return XmlElementNames.DeleteItemResponseMessage;
        }

        /// <summary>
        /// Writes XML attributes.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
        {
            base.WriteAttributesToXml(writer);

            if (this.AffectedTaskOccurrences.HasValue)
            {
                writer.WriteAttributeValue(
                    XmlAttributeNames.AffectedTaskOccurrences,
                    this.AffectedTaskOccurrences.Value);
            }

            if (this.SendCancellationsMode.HasValue)
            {
                writer.WriteAttributeValue(
                    XmlAttributeNames.SendMeetingCancellations,
                    this.SendCancellationsMode.Value);
            }

            if (this.SuppressReadReceipts)
            {
                writer.WriteAttributeValue(XmlAttributeNames.SuppressReadReceipts, true);
            }
        }

        /// <summary>
        /// Writes XML elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            this.itemIds.WriteToXml(
                writer,
                XmlNamespace.Messages,
                XmlElementNames.ItemIds);
        }

        /// <summary>
        /// Internals to json.
        /// </summary>
        /// <param name="body">The body.</param>
        protected override void InternalToJson(JsonObject body)
        {
            if (this.AffectedTaskOccurrences.HasValue)
            {
                body.Add(
                    XmlAttributeNames.AffectedTaskOccurrences,
                    this.AffectedTaskOccurrences.Value);
            }

            if (this.SendCancellationsMode.HasValue)
            {
                body.Add(
                    XmlAttributeNames.SendMeetingCancellations,
                    this.SendCancellationsMode.Value);
            }

            if (this.SuppressReadReceipts)
            {
                body.Add(XmlAttributeNames.SuppressReadReceipts, true);
            }

            if (this.ItemIds.Count > 0)
            {
                body.Add(XmlElementNames.ItemIds, this.ItemIds.InternalToJson(this.Service));
            }
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
        /// Gets the item ids.
        /// </summary>
        /// <value>The item ids.</value>
        internal ItemIdWrapperList ItemIds
        {
            get { return this.itemIds; }
        }

        /// <summary>
        /// Gets or sets the affected task occurrences.
        /// </summary>
        /// <value>The affected task occurrences.</value>
        internal AffectedTaskOccurrence? AffectedTaskOccurrences
        {
            get { return this.affectedTaskOccurrences; }
            set { this.affectedTaskOccurrences = value; }
        }

        /// <summary>
        /// Gets or sets the send cancellations.
        /// </summary>
        /// <value>The send cancellations.</value>
        internal SendCancellationsMode? SendCancellationsMode
        {
            get { return this.sendCancellationsMode; }
            set { this.sendCancellationsMode = value; }
        }

        /// <summary>
        /// Gets or sets whether to suppress read receipts
        /// </summary>
        /// <value>Whether to suppress read receipts</value>
        public bool SuppressReadReceipts
        {
            get;
            set;
        }
    }
}