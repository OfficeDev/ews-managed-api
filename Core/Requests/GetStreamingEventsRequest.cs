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
    using System.Collections.Generic;

    /// <summary>
    /// Represents a GetStreamingEvents request.
    /// </summary>
    internal class GetStreamingEventsRequest : HangingServiceRequestBase
    {
        internal const int HeartbeatFrequencyDefault = 45000; ////45s in ms
        private static int heartbeatFrequency = HeartbeatFrequencyDefault;

        private IEnumerable<string> subscriptionIds;
        private int connectionTimeout;

        /// <summary>
        /// Initializes a new instance of the <see cref="GetStreamingEventsRequest"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="serviceObjectHandler">Callback method to handle response objects received.</param>
        /// <param name="subscriptionIds">List of subscription ids to listen to on this request.</param>
        /// <param name="connectionTimeout">Connection timeout, in minutes.</param>
        internal GetStreamingEventsRequest(
            ExchangeService service, 
            HandleResponseObject serviceObjectHandler,
            IEnumerable<string> subscriptionIds,
            int connectionTimeout)
            : base(service, serviceObjectHandler, GetStreamingEventsRequest.heartbeatFrequency)
        {
            this.subscriptionIds = subscriptionIds;
            this.connectionTimeout = connectionTimeout;
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.GetStreamingEvents;
        }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.GetStreamingEventsResponse;
        }

        /// <summary>
        /// Writes the elements to XML writer.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.SubscriptionIds);

            foreach (string id in this.subscriptionIds)
            {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.SubscriptionId,
                    id);
            }

            writer.WriteEndElement();

            writer.WriteElementValue(
                XmlNamespace.Messages,
                XmlElementNames.ConnectionTimeout,
                this.connectionTimeout);
        }

        /// <summary>
        /// Gets the request version.
        /// </summary>
        /// <returns>Earliest Exchange version in which this request is supported.</returns>
        internal override ExchangeVersion GetMinimumRequiredServerVersion()
        {
            return ExchangeVersion.Exchange2010_SP1;
        }

        /// <summary>
        /// Parses the response.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>Response object.</returns>
        internal override object ParseResponse(EwsServiceXmlReader reader)
        {
            reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.ResponseMessages);

            GetStreamingEventsResponse response = new GetStreamingEventsResponse(this);
            response.LoadFromXml(reader, XmlElementNames.GetStreamingEventsResponseMessage);

            reader.ReadEndElementIfNecessary(XmlNamespace.Messages, XmlElementNames.ResponseMessages);

            return response;
        }

        #region Test hooks
        /// <summary>
        /// Allow test code to change heartbeat value
        /// </summary>
        internal static int HeartbeatFrequency
        {
            set
            {
                heartbeatFrequency = value;
            }
        }
        #endregion
    }
}