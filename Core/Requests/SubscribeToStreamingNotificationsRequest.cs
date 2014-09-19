// ---------------------------------------------------------------------------
// <copyright file="SubscribeToStreamingNotificationsRequest.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

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
