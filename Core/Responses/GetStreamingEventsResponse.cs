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
// <summary>Defines the GetStreamingEventsResponse class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents the response to a subscription event retrieval operation.
    /// </summary>
    internal sealed class GetStreamingEventsResponse : ServiceResponse
    {
        private GetStreamingEventsResults results = new GetStreamingEventsResults();
        private HangingServiceRequestBase request;

        /// <summary>
        /// Enumeration of ConnectionStatus that can be returned by the server.
        /// </summary>
        private enum ConnectionStatus
        {
            /// <summary>
            /// Simple heartbeat
            /// </summary>
            OK,

            /// <summary>
            /// Server is closing the connection.
            /// </summary>
            Closed
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="GetStreamingEventsResponse"/> class.
        /// </summary>
        /// <param name="request">Request to disconnect when we get a close message.</param>
        internal GetStreamingEventsResponse(HangingServiceRequestBase request)
            : base()
        {
            this.ErrorSubscriptionIds = new List<string>();
            this.request = request;
        }

        /// <summary>
        /// Reads response elements from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
        {
            base.ReadElementsFromXml(reader);

            reader.Read();

            if (reader.LocalName == XmlElementNames.Notifications)
            {
                this.results.LoadFromXml(reader);
            }
            else if (reader.LocalName == XmlElementNames.ConnectionStatus)
            {
                string connectionStatus = reader.ReadElementValue(XmlNamespace.Messages, XmlElementNames.ConnectionStatus);

                if (connectionStatus.Equals(ConnectionStatus.Closed.ToString()))
                {
                    this.request.Disconnect(HangingRequestDisconnectReason.Clean, null);
                }
            }
        }

        /// <summary>
        /// Loads extra error details from XML
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <param name="xmlElementName">The current element name of the extra error details.</param>
        /// <returns>
        /// True if the expected extra details is loaded;
        /// False if the element name does not match the expected element.
        /// </returns>
        internal override bool LoadExtraErrorDetailsFromXml(EwsServiceXmlReader reader, string xmlElementName)
        {
            bool baseReturnVal = base.LoadExtraErrorDetailsFromXml(reader, xmlElementName);

            if (reader.IsStartElement(XmlNamespace.Messages, XmlElementNames.ErrorSubscriptionIds))
            {
                do
                {
                    reader.Read();

                    if (reader.NodeType == System.Xml.XmlNodeType.Element &&
                        reader.LocalName == XmlElementNames.SubscriptionId)
                    {
                        this.ErrorSubscriptionIds.Add(
                            reader.ReadElementValue(XmlNamespace.Messages, XmlElementNames.SubscriptionId));
                    }
                }
                while (!reader.IsEndElement(XmlNamespace.Messages, XmlElementNames.ErrorSubscriptionIds));

                return true;
            }
            else
            {
                return baseReturnVal;
            }
        }

        /// <summary>
        /// Gets event results from subscription.
        /// </summary>
        internal GetStreamingEventsResults Results
        {
            get { return this.results; }
        }

        /// <summary>
        /// Gets the error subscription ids.
        /// </summary>
        /// <value>The error subscription ids.</value>
        internal List<string> ErrorSubscriptionIds
        {
            get;
            private set;
        }
    }
}
