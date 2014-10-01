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
// <summary>Defines the SubscribeToStreamingNotificationsRequest class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a "Streaming" Subscribe request.
    /// </summary>
    internal class SubscribeToStreamingNotificationsRequest : SubscribeRequest<StreamingSubscription>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="SubscribeToStreamingNotificationsRequest"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        internal SubscribeToStreamingNotificationsRequest(ExchangeService service)
            : base(service)
        {
        }

        /// <summary>
        /// Validate request.
        /// </summary>
        internal override void Validate()
        {
            base.Validate();

            if (!String.IsNullOrEmpty(this.Watermark))
            {
                throw new ArgumentException("Watermarks cannot be used with StreamingSubscriptions.", "Watermark");
            }
        }

        /// <summary>
        /// Gets the name of the subscription XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetSubscriptionXmlElementName()
        {
            return XmlElementNames.StreamingSubscriptionRequest;
        }

        /// <summary>
        /// Internals the write elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void InternalWriteElementsToXml(EwsServiceXmlWriter writer)
        {
        }

        /// <summary>
        /// Adds the json properties.
        /// </summary>
        /// <param name="jsonSubscribeRequest">The json subscribe request.</param>
        /// <param name="service">The service.</param>
        internal override void AddJsonProperties(JsonObject jsonSubscribeRequest, ExchangeService service)
        {
        }

        /// <summary>
        /// Creates the service response.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="responseIndex">Index of the response.</param>
        /// <returns>Service response.</returns>
        internal override SubscribeResponse<StreamingSubscription> CreateServiceResponse(ExchangeService service, int responseIndex)
        {
            return new SubscribeResponse<StreamingSubscription>(new StreamingSubscription(service));
        }

        /// <summary>
        /// Gets the request version.
        /// </summary>
        /// <returns>Earliest Exchange version in which this request is supported.</returns>
        internal override ExchangeVersion GetMinimumRequiredServerVersion()
        {
            return ExchangeVersion.Exchange2010_SP1;
        }
    }
}
